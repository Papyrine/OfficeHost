/// <summary>
/// The SQL Server stores, on LocalDB, so on Windows only.
/// </summary>
[RunOn(OS.Windows)]
[InheritsTests]
[ClassDataSource<SqlDatabase>(Shared = SharedType.PerTestSession)]
public class SqlServerTests(SqlDatabase database) :
    StorageContractTests
{
    string connectionString = database.ConnectionString;

    protected override IOfficeStorage Storage => new SqlServerStorage(connectionString);

    protected override Task<OfficeFile?> Add(string folder, string name, Stream content) =>
        new SqlServerStorage(connectionString).Add(folder, name, "owner", content);

    [Test]
    public Task Script() =>
        Verify(OfficeHostSqlServerInstaller.GetScript("office"));

    /// <summary>
    /// The tables as a Mermaid diagram, also written to <c>docs/sqlserver-tables.include.md</c>, which the readme includes.
    /// </summary>
    [Test]
    public async Task Diagram()
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync();
        await Verify(connection)
            .SchemaAsDiagram()
            .SchemaFilter(_ => _ is Microsoft.SqlServer.Management.Smo.Table {Schema: "dbo"})
            .UseFileName("SqlServerTables");
        var project = ProjectDirectory();
        // mdsnippets marks the first line of an include, which would break the fence of the diagram
        var diagram = Environment.NewLine + await File.ReadAllTextAsync(Path.Combine(project, "SqlServerTables.verified.md"));
        var include = Path.Combine(project, "..", "..", "docs", "sqlserver-tables.include.md");
        if (!File.Exists(include) ||
            await File.ReadAllTextAsync(include) != diagram)
        {
            await File.WriteAllTextAsync(include, diagram);
        }
    }

    static string ProjectDirectory([CallerFilePath] string path = "") =>
        Path.GetDirectoryName(path)!;

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
    public async Task Locks()
    {
        var time = new ManualTime();
        var locks = new SqlServerLockStore(connectionString, time);
        var fileId = folder;
        var expires = time.GetUtcNow().AddMinutes(30);
        await Assert.That(await locks.Get(fileId)).IsNull();
        await Assert.That(await locks.CompareExchange(fileId, null, new("a", expires))).IsNull();
        await Assert.That(await locks.Get(fileId)).IsEqualTo(new("a", expires));

        // a mismatch changes nothing
        await Assert.That(await locks.CompareExchange(fileId, null, new("b", expires))).IsEqualTo(new("a", expires));
        await Assert.That(await locks.CompareExchange(fileId, "b", null)).IsEqualTo(new("a", expires));
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("a");

        await Assert.That(await locks.CompareExchange(fileId, "a", new("b", expires))).IsEqualTo(new("a", expires));
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("b");
        await Assert.That(await locks.CompareExchange(fileId, "b", null)).IsEqualTo(new("b", expires));
        await Assert.That(await locks.Get(fileId)).IsNull();

        // an expired lock is no lock
        await locks.CompareExchange(fileId, null, new("c", expires));
        time.Advance(TimeSpan.FromMinutes(31));
        await Assert.That(await locks.Get(fileId)).IsNull();
        await Assert.That(await locks.CompareExchange(fileId, null, new("d", time.GetUtcNow().AddMinutes(30)))).IsNull();
        await Assert.That((await locks.Get(fileId))!.Id).IsEqualTo("d");
    }

    /// <summary>
    /// Lock ids are compared as they are, whatever the collation of the database, and a comparand longer than a lock id
    /// is not cut into one.
    /// </summary>
    [Test]
    public async Task Locks_ComparedAsTheyAre()
    {
        var locks = new SqlServerLockStore(connectionString, TimeProvider.System);
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        await locks.CompareExchange(folder, null, new("ABC", expires));
        await Assert.That((await locks.CompareExchange(folder, "abc", null))!.Id).IsEqualTo("ABC");
        await Assert.That((await locks.Get(folder))!.Id).IsEqualTo("ABC");

        var id = new string('a', 1024);
        await locks.CompareExchange(folder, "ABC", new(id, expires));
        await Assert.That((await locks.CompareExchange(folder, id + "b", null))!.Id).IsEqualTo(id);
        await Assert.That((await locks.Get(folder))!.Id).IsEqualTo(id);
    }

    /// <summary>
    /// Tables of version 1, from before the version table, are brought up to date, keeping their rows: their ids are
    /// then compared as they are.
    /// </summary>
    [Test]
    public async Task Install_FromVersion1()
    {
        var schema = $"v1{Guid.NewGuid():N}";
        await using (var connection = new SqlConnection(connectionString))
        {
            await connection.OpenAsync();
            await using var create = new SqlCommand(
                $"""
                 exec('create schema [{schema}]');
                 create table [{schema}].[OfficeHostFiles] (Id nvarchar(450) not null primary key, Folder nvarchar(400) not null, Name nvarchar(255) not null, OwnerId nvarchar(450) not null, Content varbinary(max) not null, Size as datalength(Content), Version uniqueidentifier not null, LastModified datetimeoffset not null);
                 create unique index IX_OfficeHostFiles_Folder_Name on [{schema}].[OfficeHostFiles] (Folder, Name);
                 create table [{schema}].[OfficeHostLocks] (FileId nvarchar(450) not null primary key, LockId nvarchar(1024) not null, Expires datetimeoffset not null);
                 create table [{schema}].[OfficeHostFssHttpStates] (FileId nvarchar(450) not null, Partition uniqueidentifier not null, Version nvarchar(max) not null, Data varbinary(max) not null, primary key (FileId, Partition));
                 create table [{schema}].[OfficeHostUserInfo] (UserId nvarchar(450) not null primary key, Info nvarchar(max) not null);
                 """,
                connection);
            await create.ExecuteNonQueryAsync();
        }

        var storage = new SqlServerStorage(connectionString, schema);
        var file = (await storage.Add("folder", "kept.docx", "owner", Content("kept")))!;
        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("kept");

        var locks = new SqlServerLockStore(connectionString, TimeProvider.System, schema);
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        await locks.CompareExchange("Doc", null, new("upper", expires));
        await Assert.That(await locks.CompareExchange("doc", null, new("lower", expires))).IsNull();
        await Assert.That((await locks.Get("Doc"))!.Id).IsEqualTo("upper");
    }

    /// <summary>
    /// The content is read in chunks: a file of several, and a write while it is read, which fails the read rather than
    /// mixing two versions.
    /// </summary>
    [Test]
    public async Task OpenRead_Chunks()
    {
        var storage = new SqlServerStorage(connectionString);
        var bytes = RandomNumberGenerator.GetBytes(9 * 1024 * 1024 + 7);
        var file = (await storage.Add(folder, "large.bin", "owner", new MemoryStream(bytes)))!;
        await using (var read = await storage.OpenRead(file.Id))
        {
            using var copy = new MemoryStream();
            await read.CopyToAsync(copy);
            await Assert.That(copy.ToArray().AsSpan().SequenceEqual(bytes)).IsTrue();
        }

        await using var interrupted = await storage.OpenRead(file.Id);
        var buffer = new byte[1024];
        await interrupted.ReadExactlyAsync(buffer);
        await storage.Write(file.Id, new MemoryStream(bytes));
        await Assert.That(() => interrupted.CopyToAsync(Stream.Null)).Throws<IOException>();
    }

    /// <summary>
    /// A name longer than the table holds is refused, not cut.
    /// </summary>
    [Test]
    public async Task Add_NameTooLong()
    {
        var storage = new SqlServerStorage(connectionString);
        await Assert.That(() => storage.Add(folder, new('n', 256), "owner", Content("x"))).Throws<ArgumentException>();
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
