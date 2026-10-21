using Testcontainers.Azurite;
using TUnit.Core.Interfaces;

/// <summary>
/// The blob service of the test session: a real storage account when <c>OFFICEHOST_BLOB_CONNECTION_STRING</c> is set,
/// otherwise Azurite in Docker. Each test class uses a container of its own, deleted at the end.
/// </summary>
public sealed class BlobAccount :
    IAsyncInitializer,
    IAsyncDisposable
{
    AzuriteContainer? azurite;
    List<BlobContainerClient> containers = [];

    public BlobServiceClient Service { get; private set; } = null!;

    public async Task InitializeAsync()
    {
        var connectionString = Environment.GetEnvironmentVariable("OFFICEHOST_BLOB_CONNECTION_STRING");
        if (connectionString == null)
        {
            azurite = new AzuriteBuilder("mcr.microsoft.com/azure-storage/azurite:3.35.0")
                .WithInMemoryPersistence()
                // the Azure SDK is often ahead of Azurite
                .WithCommand("--skipApiVersionCheck")
                .Build();
            await azurite.StartAsync();
            connectionString = azurite.GetConnectionString();
        }

        Service = new(connectionString);
    }

    public async Task<BlobContainerClient> NewContainer()
    {
        var container = Service.GetBlobContainerClient($"officehost-tests-{Guid.NewGuid():N}");
        await container.CreateAsync();
        lock (containers)
        {
            containers.Add(container);
        }

        return container;
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var container in containers)
        {
            await container.DeleteIfExistsAsync();
        }

        if (azurite != null)
        {
            await azurite.DisposeAsync();
        }
    }
}
