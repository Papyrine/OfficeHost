namespace OfficeHost;

/// <summary>
/// Creates the tables of the SQL Server stores of OfficeHost.SqlServer, and their schema, when missing, and brings
/// tables of an older version up to date: the version is kept in a table of its own. Running it again changes nothing,
/// and two instances that run it at once take turns.
/// </summary>
public static class OfficeHostSqlServerInstaller
{
    /// <summary>
    /// The version of the tables the script makes: 1, the tables; 2, ids, lock ids and versions compared as they are, in
    /// a binary collation, whatever that of the database.
    /// </summary>
    public const int Version = 2;

    const string binary = "Latin1_General_100_BIN2";

    public static async Task Install(string connectionString, string schema = "dbo", Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await Install(connection, schema, cancel);
    }

    /// <param name="connection">An open connection.</param>
    public static async Task Install(SqlConnection connection, string schema = "dbo", Cancel cancel = default)
    {
        // the script is a transaction of its own
        await using var command = new SqlCommand(GetScript(schema), connection);
        command.CommandTimeout = 0;
        await command.ExecuteNonQueryAsync(cancel);
    }

    /// <summary>
    /// The script that <see cref="Install(string, string, Cancel)"/> runs, to apply by hand.
    /// </summary>
    public static string GetScript(string schema = "dbo")
    {
        var quoted = SqlServerTables.Quote(schema);
        var files = SqlServerTables.Name(schema, SqlServerTables.Files);
        var locks = SqlServerTables.Name(schema, SqlServerTables.Locks);
        var states = SqlServerTables.Name(schema, SqlServerTables.FssHttpStates);
        var userInfo = SqlServerTables.Name(schema, SqlServerTables.UserInfo);
        var versions = SqlServerTables.Name(schema, SqlServerTables.SchemaVersion);
        return $"""
                set xact_abort on;
                begin transaction;
                -- installs in the database take turns, whatever their schema: the DDL of two at once deadlocks on the
                -- system catalog
                exec sp_getapplock @Resource = N'OfficeHost.Install', @LockMode = 'Exclusive', @LockOwner = 'Transaction';

                if schema_id({Literal(schema)}) is null
                    exec({Literal($"create schema {quoted}")});

                if object_id({Literal(versions)}) is null
                    create table {versions} (Version int not null);

                declare @version int = (select max(Version) from {versions});
                -- the tables of before the version table are version 1
                if @version is null and object_id({Literal(files)}) is not null
                    set @version = 1;

                -- new tables are made as they are now; older ones are brought up to date
                if @version is null
                begin
                {Tables(files, locks, states, userInfo)}
                end;

                if @version = 1
                begin
                {Version2(files, locks, states, userInfo)}
                end;

                delete from {versions};
                insert into {versions} (Version) values ({Version});
                commit;
                """;
    }

    /// <summary>
    /// The tables of the current version, for a new install: what version 1 then the changes of the later versions make,
    /// without changing a table once made.
    /// </summary>
    static string Tables(string files, string locks, string states, string userInfo) =>
        $"""
         if object_id({Literal(files)}) is null
         begin
             create table {files} (
                 Id nvarchar(450) collate {binary} not null constraint {Key(SqlServerTables.Files)} primary key,
                 Folder nvarchar(400) not null,
                 Name nvarchar(255) not null,
                 OwnerId nvarchar(450) not null,
                 Content varbinary(max) not null,
                 Size as datalength(Content),
                 Version uniqueidentifier not null,
                 LastModified datetimeoffset not null);
             create unique index IX_{SqlServerTables.Files}_Folder_Name on {files} (Folder, Name);
         end;

         if object_id({Literal(locks)}) is null
             create table {locks} (
                 FileId nvarchar(450) collate {binary} not null constraint {Key(SqlServerTables.Locks)} primary key,
                 LockId nvarchar(1024) collate {binary} not null,
                 Expires datetimeoffset not null);

         if object_id({Literal(states)}) is null
             create table {states} (
                 FileId nvarchar(450) collate {binary} not null,
                 Partition uniqueidentifier not null,
                 Version nvarchar(max) collate {binary} not null,
                 Data varbinary(max) not null,
                 constraint {Key(SqlServerTables.FssHttpStates)} primary key (FileId, Partition));

         if object_id({Literal(userInfo)}) is null
             create table {userInfo} (
                 UserId nvarchar(450) collate {binary} not null constraint {Key(SqlServerTables.UserInfo)} primary key,
                 Info nvarchar(max) not null);
         """;

    static string Key(string tableName) =>
        SqlServerTables.Quote($"PK_{tableName}");

    /// <summary>
    /// Version 2 from version 1: the ids, lock ids and versions in a binary collation. With that of the database,
    /// usually case insensitive, the lock <c>ABC</c> was released by <c>abc</c>, and the ids <c>Doc</c> and <c>doc</c>
    /// shared a lock.
    /// </summary>
    static string Version2(string files, string locks, string states, string userInfo) =>
        $"""
         {Rekey(SqlServerTables.Files, files, "Id nvarchar(450)", "Id")}
         {Rekey(SqlServerTables.Locks, locks, "FileId nvarchar(450)", "FileId")}
         exec({Literal($"alter table {locks} alter column LockId nvarchar(1024) collate {binary} not null")});
         {Rekey(SqlServerTables.FssHttpStates, states, "FileId nvarchar(450)", "FileId, Partition")}
         exec({Literal($"alter table {states} alter column Version nvarchar(max) collate {binary} not null")});
         {Rekey(SqlServerTables.UserInfo, userInfo, "UserId nvarchar(450)", "UserId")}
         """;

    /// <summary>
    /// Changes the collation of a column of the primary key: the key is dropped, by the name SQL Server gave it, and
    /// made again under a name of its own.
    /// </summary>
    static string Rekey(string tableName, string table, string column, string key)
    {
        var variable = $"@{tableName}Drop";
        return $"""
                declare {variable} nvarchar(max) = {Literal($"alter table {table} drop constraint ")} + quotename((select name from sys.key_constraints where parent_object_id = object_id({Literal(table)}) and type = 'PK'));
                exec({variable});
                exec({Literal($"alter table {table} alter column {column} collate {binary} not null")});
                exec({Literal($"alter table {table} add constraint {Key(tableName)} primary key ({key})")});
                """;
    }

    static string Literal(string value) =>
        $"N'{value.Replace("'", "''")}'";
}
