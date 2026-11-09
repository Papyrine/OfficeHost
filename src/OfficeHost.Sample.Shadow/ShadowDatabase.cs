using Microsoft.Data.SqlClient;

/// <summary>
/// The database of the shadow: created when missing, then the tables of OfficeHost.SqlServer in it. The account needs
/// to create a database, or the database must exist; and to create tables, or <c>Shadow:InstallDatabase</c> is false
/// and the tables are made with the script of <see cref="OfficeHostSqlServerInstaller.GetScript"/>.
/// </summary>
static class ShadowDatabase
{
    public static async Task Install(string connectionString, string schema)
    {
        var builder = new SqlConnectionStringBuilder(connectionString);
        var database = builder.InitialCatalog;
        if (database.Length > 0)
        {
            builder.InitialCatalog = "master";
            await using var connection = new SqlConnection(builder.ConnectionString);
            await connection.OpenAsync();
            await using var command = new SqlCommand(
                """
                if db_id(@name) is null
                    exec('create database ' + @quoted);
                """,
                connection);
            command.Parameters.AddWithValue("@name", database);
            command.Parameters.AddWithValue("@quoted", $"[{database.Replace("]", "]]")}]");
            await command.ExecuteNonQueryAsync();
        }

        await OfficeHostSqlServerInstaller.Install(connectionString, schema);
    }
}
