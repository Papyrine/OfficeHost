/// <summary>
/// The installer, on a database of its own, one test at a time: DDL takes locks on the system catalog of the database,
/// in which the queries of the store tests would wait, or deadlock.
/// </summary>
[RunOn(OS.Windows)]
[NotInParallel]
[ClassDataSource<SqlDatabase>(Shared = SharedType.PerClass)]
public class SqlServerInstallTests(SqlDatabase database)
{
    string connectionString = database.ConnectionString;

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

    /// <summary>
    /// Installing again changes nothing, and a new install has the tables of the current version: ids compared as they
    /// are.
    /// </summary>
    [Test]
    public async Task Install_Again()
    {
        await OfficeHostSqlServerInstaller.Install(connectionString);
        await OfficeHostSqlServerInstaller.Install(connectionString, "other");
        await OfficeHostSqlServerInstaller.Install(connectionString, "other");
        var locks = new SqlServerLockStore(connectionString, TimeProvider.System, "other");
        await Assert.That(await locks.Get("file")).IsNull();
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        await locks.CompareExchange("Doc", null, new("upper", expires));
        await Assert.That(await locks.CompareExchange("doc", null, new("lower", expires))).IsNull();
    }

    /// <summary>
    /// Installs at once, in schemas of their own, take turns rather than deadlock.
    /// </summary>
    [Test]
    public Task Install_AtOnce() =>
        Task.WhenAll(
            Enumerable.Range(0, 4)
                .Select(_ => Task.Run(() => OfficeHostSqlServerInstaller.Install(connectionString, $"atonce{_}"))));

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
        var file = await storage.Add("folder", "kept.docx", "owner", new MemoryStream("kept"u8.ToArray()));
        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        await using (var stream = await storage.OpenRead(file.Id))
        {
            using var reader = new StreamReader(stream);
            await Assert.That(await reader.ReadToEndAsync()).IsEqualTo("kept");
        }

        var locks = new SqlServerLockStore(connectionString, TimeProvider.System, schema);
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        await locks.CompareExchange("Doc", null, new("upper", expires));
        await Assert.That(await locks.CompareExchange("doc", null, new("lower", expires), new OfficeUser("bob", "Bob", OfficeClient.Desktop))).IsNull();
        await Assert.That((await locks.Get("Doc"))!.Id).IsEqualTo("upper");
        await Assert.That((await locks.Get("doc"))!.Holder).IsEqualTo(new OfficeUser("bob", "Bob", OfficeClient.Desktop));
    }

    /// <summary>
    /// Tables of version 2 get the holder of a lock, keeping their rows: the locks of before have none.
    /// </summary>
    [Test]
    public async Task Install_FromVersion2()
    {
        var schema = $"v2{Guid.NewGuid():N}";
        await using (var connection = new SqlConnection(connectionString))
        {
            await connection.OpenAsync();
            await using var create = new SqlCommand(
                $"""
                 exec('create schema [{schema}]');
                 create table [{schema}].[OfficeHostFiles] (Id nvarchar(450) collate Latin1_General_100_BIN2 not null constraint [PK_OfficeHostFiles] primary key, Folder nvarchar(400) not null, Name nvarchar(255) not null, OwnerId nvarchar(450) not null, Content varbinary(max) not null, Size as datalength(Content), Version uniqueidentifier not null, LastModified datetimeoffset not null);
                 create unique index IX_OfficeHostFiles_Folder_Name on [{schema}].[OfficeHostFiles] (Folder, Name);
                 create table [{schema}].[OfficeHostLocks] (FileId nvarchar(450) collate Latin1_General_100_BIN2 not null constraint [PK_OfficeHostLocks] primary key, LockId nvarchar(1024) collate Latin1_General_100_BIN2 not null, Expires datetimeoffset not null);
                 create table [{schema}].[OfficeHostFssHttpStates] (FileId nvarchar(450) collate Latin1_General_100_BIN2 not null, Partition uniqueidentifier not null, Version nvarchar(max) collate Latin1_General_100_BIN2 not null, Data varbinary(max) not null, constraint [PK_OfficeHostFssHttpStates] primary key (FileId, Partition));
                 create table [{schema}].[OfficeHostUserInfo] (UserId nvarchar(450) collate Latin1_General_100_BIN2 not null constraint [PK_OfficeHostUserInfo] primary key, Info nvarchar(max) not null);
                 create table [{schema}].[OfficeHostSchemaVersion] (Version int not null);
                 insert into [{schema}].[OfficeHostSchemaVersion] (Version) values (2);
                 insert into [{schema}].[OfficeHostLocks] (FileId, LockId, Expires) values ('kept', 'L', '2100-01-01');
                 """,
                connection);
            await create.ExecuteNonQueryAsync();
        }

        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
        var locks = new SqlServerLockStore(connectionString, TimeProvider.System, schema);
        await Assert.That(await locks.Get("kept")).IsEqualTo(new OfficeLock("L", new(2100, 1, 1, 0, 0, 0, TimeSpan.Zero)));
        var alice = new OfficeUser("alice", "Alice", OfficeClient.Wopi);
        await Assert.That(await locks.CompareExchange("new", null, new("M", DateTimeOffset.UtcNow.AddMinutes(30)), alice)).IsNull();
        await Assert.That((await locks.Get("new"))!.Holder).IsEqualTo(alice);
    }
}
