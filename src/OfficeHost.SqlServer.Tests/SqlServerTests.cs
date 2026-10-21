/// <summary>
/// The SQL Server stores, on LocalDB, so on Windows only.
/// </summary>
[RunOn(OS.Windows)]
[ClassDataSource<SqlDatabase>(Shared = SharedType.PerTestSession)]
public class SqlServerTests(SqlDatabase database)
{
    string connectionString = database.ConnectionString;

    // each test has a folder and files of its own
    string folder = Guid.NewGuid().ToString("N");

    static MemoryStream Content(string content) =>
        new(Encoding.UTF8.GetBytes(content));

    static async Task<string> Read(Stream stream)
    {
        await using (stream)
        {
            using var reader = new StreamReader(stream);
            return await reader.ReadToEndAsync();
        }
    }

    [Test]
    public Task Script() =>
        Verify(OfficeHostSqlServerInstaller.GetScript("office"));

    [Test]
    public async Task Install_Again()
    {
        await OfficeHostSqlServerInstaller.Install(connectionString);
        await OfficeHostSqlServerInstaller.Install(connectionString, "other");
        await OfficeHostSqlServerInstaller.Install(connectionString, "other");
        var locks = new SqlServerLockStore(connectionString, TimeProvider.System, "other");
        await Assert.That(await locks.Get("file")).IsNull();
    }

    [Test]
    public async Task Storage()
    {
        var storage = new SqlServerStorage(connectionString);
        var file = (await storage.Add(folder, "test.docx", "owner", Content("12345")))!;
        await Assert.That(file.Name).IsEqualTo("test.docx");
        await Assert.That(file.Size).IsEqualTo(5);
        await Assert.That(file.OwnerId).IsEqualTo("owner");
        await Assert.That(await storage.GetFile(file.Id)).IsEqualTo(file);
        await Assert.That(await storage.Add(folder, "TEST.docx", "owner", Content(""))).IsNull();
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("12345");

        var written = await storage.Write(file.Id, Content("changed"));
        await Assert.That(written.Id).IsEqualTo(file.Id);
        await Assert.That(written.Size).IsEqualTo(7);
        await Assert.That(written.Version).IsNotEqualTo(file.Version);
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("changed");

        await Assert.That(await storage.GetSibling(file.Id, "other.docx")).IsNull();
        var sibling = (await storage.CreateSibling(file.Id, "other.docx", Content("other"), false))!;
        await Assert.That(sibling.Id).IsNotEqualTo(file.Id);
        await Assert.That(sibling.OwnerId).IsEqualTo("owner");
        await Assert.That(await storage.GetSibling(file.Id, "other.docx")).IsEqualTo(sibling);
        await Assert.That(await storage.CreateSibling(file.Id, "other.docx", Content("again"), false)).IsNull();
        var overwritten = (await storage.CreateSibling(file.Id, "other.docx", Content("again"), true))!;
        await Assert.That(overwritten.Id).IsEqualTo(sibling.Id);
        await Assert.That(await Read(await storage.OpenRead(sibling.Id))).IsEqualTo("again");

        // a sibling is in the same folder only
        var elsewhere = (await storage.Add($"{folder}/sub", "other.docx", "owner", Content("")))!;
        await Assert.That(await storage.GetSibling(elsewhere.Id, "test.docx")).IsNull();

        await Assert.That(await storage.Rename(file.Id, "other.docx")).IsNull();
        var renamed = (await storage.Rename(file.Id, "renamed.docx"))!;
        await Assert.That(renamed.Id).IsEqualTo(file.Id);
        await Assert.That(renamed.Name).IsEqualTo("renamed.docx");
        // a change of case only is the same file
        await Assert.That((await storage.Rename(file.Id, "Renamed.docx"))!.Name).IsEqualTo("Renamed.docx");

        await storage.Delete(file.Id);
        await Assert.That(await storage.GetFile(file.Id)).IsNull();
    }

    [Test]
    public async Task Storage_Missing()
    {
        var storage = new SqlServerStorage(connectionString);
        await Assert.That(await storage.GetFile("missing")).IsNull();
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.OpenRead("missing"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Write("missing", Content("")));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.GetSibling("missing", "a.docx"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.CreateSibling("missing", "a.docx", Content(""), true));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Rename("missing", "a.docx"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Delete("missing"));
    }

    [Test]
    public async Task Storage_Large()
    {
        var storage = new SqlServerStorage(connectionString);
        var bytes = RandomNumberGenerator.GetBytes(20 * 1024 * 1024);
        var file = (await storage.Add(folder, "large.docx", "owner", new MemoryStream(bytes)))!;
        await Assert.That(file.Size).IsEqualTo(bytes.Length);
        await using var stream = await storage.OpenRead(file.Id);
        var read = new MemoryStream();
        await stream.CopyToAsync(read);
        await Assert.That(read.ToArray().SequenceEqual(bytes)).IsTrue();
    }

    [Test]
    public async Task Locks()
    {
        var time = new ManualTime();
        var locks = new SqlServerLockStore(connectionString, time);
        var fileId = folder;
        var expires = time.GetUtcNow().AddMinutes(30);
        await Assert.That(await locks.Get(fileId)).IsNull();
        await Assert.That(await locks.CompareExchange(fileId, null, new("a", expires))).IsNull();
        await Assert.That(await locks.Get(fileId)).IsEqualTo(new OfficeLock("a", expires));

        // a mismatch changes nothing
        await Assert.That(await locks.CompareExchange(fileId, null, new("b", expires))).IsEqualTo(new OfficeLock("a", expires));
        await Assert.That(await locks.CompareExchange(fileId, "b", null)).IsEqualTo(new OfficeLock("a", expires));
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("a");

        await Assert.That(await locks.CompareExchange(fileId, "a", new("b", expires))).IsEqualTo(new OfficeLock("a", expires));
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("b");
        await Assert.That(await locks.CompareExchange(fileId, "b", null)).IsEqualTo(new OfficeLock("b", expires));
        await Assert.That(await locks.Get(fileId)).IsNull();

        // an expired lock is no lock
        await locks.CompareExchange(fileId, null, new("c", expires));
        time.Advance(TimeSpan.FromMinutes(31));
        await Assert.That(await locks.Get(fileId)).IsNull();
        await Assert.That(await locks.CompareExchange(fileId, null, new("d", time.GetUtcNow().AddMinutes(30)))).IsNull();
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("d");
    }

    [Test]
    public async Task Locks_Race()
    {
        var locks = new SqlServerLockStore(connectionString, TimeProvider.System);
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        var results = await Task.WhenAll(
            Enumerable.Range(0, 20)
                .Select(_ => Task.Run(() => locks.CompareExchange(folder, null, new($"{_}", expires)))));
        await Assert.That(results.Count(_ => _ == null)).IsEqualTo(1);
    }

    [Test]
    public async Task FssHttp()
    {
        var store = new SqlServerFssHttpStore(connectionString);
        var partition = Guid.NewGuid();
        await Assert.That(await store.Get(folder, partition)).IsNull();
        await Assert.That(await store.CompareExchange(folder, partition, "x", new("1", new byte[] {1}))).IsFalse();
        await Assert.That(await store.CompareExchange(folder, partition, null, new("1", new byte[] {1}))).IsTrue();
        await Assert.That(await store.CompareExchange(folder, partition, null, new("2", new byte[] {2}))).IsFalse();
        await Assert.That(await store.CompareExchange(folder, partition, "1", new("2", new byte[] {2, 3}))).IsTrue();
        var state = (await store.Get(folder, partition))!;
        await Assert.That(state.Version).IsEqualTo("2");
        await Assert.That(state.Data.ToArray()).IsEquivalentTo(new byte[] {2, 3});
        await Assert.That(await store.Get(folder, Guid.NewGuid())).IsNull();

        await store.Delete(folder);
        await Assert.That(await store.Get(folder, partition)).IsNull();
    }

    [Test]
    public async Task FssHttp_Race()
    {
        var store = new SqlServerFssHttpStore(connectionString);
        var partition = Guid.NewGuid();
        var results = await Task.WhenAll(
            Enumerable.Range(0, 20)
                .Select(_ => Task.Run(() => store.CompareExchange(folder, partition, null, new($"{_}", new byte[] {1})))));
        await Assert.That(results.Count(_ => _)).IsEqualTo(1);
    }

    [Test]
    public async Task UserInfo()
    {
        var store = new SqlServerUserInfoStore(connectionString);
        await Assert.That(await store.Get(folder)).IsNull();
        await store.Set(folder, "first");
        await store.Set(folder, "second");
        await Assert.That(await store.Get(folder)).IsEqualTo("second");
    }

    /// <summary>
    /// The WOPI endpoints over the stores of <see cref="OfficeHostSqlServerExtensions.AddOfficeHostSqlServer"/>.
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
        // begin-snippet: SqlServer
        // once, from a deployment step, with rights to create tables
        await OfficeHostSqlServerInstaller.Install(connectionString);

        // in the application, before or after AddOfficeHost
        builder.Services.AddOfficeHostSqlServer(connectionString);
        // end-snippet
        await using var app = builder.Build();
        app.MapOfficeHost();
        await app.StartAsync();

        await Assert.That(app.Services.GetRequiredService<IOfficeLockStore>()).IsTypeOf<SqlServerLockStore>();
        var storage = app.Services.GetRequiredService<SqlServerStorage>();
        var file = (await storage.Add(folder, "test.docx", "owner", Content("12345")))!;
        var token = await app.Services.GetRequiredService<IOfficeTokens>()
            .Issue(new(file.Id, "user") { CanWrite = true });
        using var client = app.GetTestClient();
        var url = $"/wopi/files/{file.Id}?access_token={token.Value}";

        using var info = await client.GetAsync(url);
        await Assert.That(info.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await info.Content.ReadAsStringAsync()).Contains("\"Size\":5");

        using var lockRequest = new HttpRequestMessage(HttpMethod.Post, url);
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
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("saved");

        using var contents = await client.GetAsync($"/wopi/files/{file.Id}/contents?access_token={token.Value}");
        await Assert.That(await contents.Content.ReadAsStringAsync()).IsEqualTo("saved");
        await app.StopAsync();
    }
}
