namespace OfficeHost;

public static class OfficeHostAzureBlobExtensions
{
    /// <summary>
    /// Registers <see cref="BlobStorage"/> as the <see cref="IOfficeStorage"/>, over a container that exists.
    /// </summary>
    public static IServiceCollection AddOfficeHostBlobStorage(this IServiceCollection services, BlobContainerClient container)
    {
        services.AddSingleton(new BlobStorage(container));
        services.AddSingleton<IOfficeStorage>(_ => _.GetRequiredService<BlobStorage>());
        return services;
    }
}
