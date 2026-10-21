using TUnit.Core.Interfaces;

/// <summary>
/// A database of its own on LocalDB for the test session, with the tables installed, dropped at the end.
/// </summary>
public sealed class SqlDatabase :
    IAsyncInitializer,
    IAsyncDisposable
{
    const string server = @"Server=(localdb)\MSSQLLocalDB;Integrated Security=true;TrustServerCertificate=true";

    string name = $"OfficeHostTests_{Guid.NewGuid():N}";

    public string ConnectionString => $"{server};Database={name}";

    public async Task InitializeAsync()
    {
        await Execute($"create database [{name}]");
        await OfficeHostSqlServerInstaller.Install(ConnectionString);
    }

    public async ValueTask DisposeAsync()
    {
        SqlConnection.ClearAllPools();
        await Execute($"alter database [{name}] set single_user with rollback immediate; drop database [{name}]");
    }

    static async Task Execute(string sql)
    {
        await using var connection = new SqlConnection(server);
        await connection.OpenAsync();
        await using var command = new SqlCommand(sql, connection);
        await command.ExecuteNonQueryAsync();
    }
}
