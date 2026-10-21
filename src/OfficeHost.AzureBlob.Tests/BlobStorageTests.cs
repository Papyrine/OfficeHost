/// <summary>
/// <see cref="BlobStorage"/>, on Azurite or a storage account, see <see cref="BlobAccount"/>.
/// </summary>
[InheritsTests]
[ClassDataSource<BlobAccount>(Shared = SharedType.PerTestSession)]
public class BlobStorageTests(BlobAccount account) :
    StorageContractTests
{
    BlobContainerClient container = null!;
    BlobStorage storage = null!;

    [Before(Test)]
    public async Task CreateContainer()
    {
        container = await account.NewContainer();
        storage = new(container);
    }

    protected override IOfficeStorage Storage => storage;

    protected override Task<OfficeFile?> Add(string folder, string name, Stream content) =>
        storage.Add(folder, name, "owner", content);

    /// <summary>
    /// A name whose file is gone, as left by a crash between deleting the file and its name, is taken over.
    /// </summary>
    [Test]
    public async Task LeftoverName()
    {
        var file = (await Add(folder, "test.docx", Content("")))!;
        await container.GetBlobClient($"files/{file.Id}").DeleteAsync();
        await Assert.That(await Add(folder, "test.docx", Content(""))).IsNotNull();
    }

    /// <summary>
    /// The WOPI endpoints over <see cref="OfficeHostAzureBlobExtensions.AddOfficeHostBlobStorage"/>.
    /// </summary>
    [Test]
    public async Task Wopi()
    {
        var builder = WebApplication.CreateSlimBuilder();
        builder.WebHost.UseTestServer();
        builder.Logging.ClearProviders();
        builder.Services.AddOfficeHost(_ =>
        {
            _.SigningKey = "the signing key of the tests, long enough";
            _.BaseUrl = "http://localhost";
        });
        builder.Services.AddOfficeHostBlobStorage(await account.NewContainer());
        await using var app = builder.Build();
        app.MapOfficeHost();
        await app.StartAsync();

        var blobs = app.Services.GetRequiredService<BlobStorage>();
        var file = (await blobs.Add(folder, "test.docx", "owner", Content("12345")))!;
        var token = await app.Services.GetRequiredService<IOfficeTokens>()
            .Issue(new(file.Id, "user") { CanWrite = true });
        using var client = app.GetTestClient();

        using var info = await client.GetAsync($"/wopi/files/{file.Id}?access_token={token.Value}");
        await Assert.That(info.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await info.Content.ReadAsStringAsync()).Contains("\"Size\":5");

        using var lockRequest = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{file.Id}?access_token={token.Value}");
        lockRequest.Headers.Add("X-WOPI-Override", "LOCK");
        lockRequest.Headers.Add("X-WOPI-Lock", "lock");
        using var locked = await client.SendAsync(lockRequest);
        await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.OK);

        using var put = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{file.Id}/contents?access_token={token.Value}")
        {
            Content = new StringContent("saved")
        };
        put.Headers.Add("X-WOPI-Override", "PUT");
        put.Headers.Add("X-WOPI-Lock", "lock");
        using var saved = await client.SendAsync(put);
        await Assert.That(saved.StatusCode).IsEqualTo(HttpStatusCode.OK);

        using var contents = await client.GetAsync($"/wopi/files/{file.Id}/contents?access_token={token.Value}");
        await Assert.That(await contents.Content.ReadAsStringAsync()).IsEqualTo("saved");
        await app.StopAsync();
    }
}
