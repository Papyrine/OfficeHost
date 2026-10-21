namespace OfficeHost;

/// <summary>
/// Creates the tables of the SQL Server stores of OfficeHost.SqlServer, and their schema, when missing. Running it again
/// changes nothing.
/// </summary>
public static class OfficeHostSqlServerInstaller
{
    public static async Task Install(string connectionString, string schema = "dbo", Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await Install(connection, schema, cancel);
    }

    /// <param name="connection">An open connection.</param>
    public static async Task Install(SqlConnection connection, string schema = "dbo", Cancel cancel = default)
    {
        await using var transaction = (SqlTransaction) await connection.BeginTransactionAsync(cancel);
        await using var command = new SqlCommand(GetScript(schema), connection, transaction);
        await command.ExecuteNonQueryAsync(cancel);
        await transaction.CommitAsync(cancel);
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
        return $"""
                if schema_id({Literal(schema)}) is null
                    exec({Literal($"create schema {quoted}")});

                if object_id({Literal(files)}) is null
                begin
                    create table {files} (
                        Id nvarchar(450) not null primary key,
                        Folder nvarchar(400) not null,
                        Name nvarchar(255) not null,
                        OwnerId nvarchar(450) not null,
                        Content varbinary(max) not null,
                        Size as datalength(Content),
                        Version rowversion not null,
                        LastModified datetimeoffset not null);
                    create unique index IX_{SqlServerTables.Files}_Folder_Name on {files} (Folder, Name);
                end;

                if object_id({Literal(locks)}) is null
                    create table {locks} (
                        FileId nvarchar(450) not null primary key,
                        LockId nvarchar(1024) not null,
                        Expires datetimeoffset not null);

                if object_id({Literal(states)}) is null
                    create table {states} (
                        FileId nvarchar(450) not null,
                        Partition uniqueidentifier not null,
                        Version nvarchar(max) not null,
                        Data varbinary(max) not null,
                        primary key (FileId, Partition));

                if object_id({Literal(userInfo)}) is null
                    create table {userInfo} (
                        UserId nvarchar(450) not null primary key,
                        Info nvarchar(max) not null);
                """;
    }

    static string Literal(string value) =>
        $"N'{value.Replace("'", "''")}'";
}
