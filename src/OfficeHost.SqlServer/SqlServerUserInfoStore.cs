namespace OfficeHost;

/// <summary>
/// The user info of PutUserInfo in a SQL Server table.
/// </summary>
public sealed class SqlServerUserInfoStore(string connectionString, string schema = "dbo") :
    IOfficeUserInfoStore
{
    string table = SqlServerTables.Name(schema, SqlServerTables.UserInfo);

    public Task<string?> Get(string userId, Cancel cancel = default) =>
        DeadlockRetry.Once(() => GetAttempt(userId, cancel));

    async Task<string?> GetAttempt(string userId, Cancel cancel)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand($"select Info from {table} where UserId = @userId", connection);
        command.Parameters.AddWithValue("@userId", userId);
        return (string?) await command.ExecuteScalarAsync(cancel);
    }

    public Task Set(string userId, string userInfo, Cancel cancel = default) =>
        DeadlockRetry.Once(() => SetAttempt(userId, userInfo, cancel));

    async Task SetAttempt(string userId, string userInfo, Cancel cancel)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand(
            $"""
             set xact_abort on;
             begin transaction;
             update {table} with (updlock, serializable) set Info = @info where UserId = @userId;
             if @@rowcount = 0
                 insert into {table} (UserId, Info) values (@userId, @info);
             commit;
             """,
            connection);
        command.Parameters.AddWithValue("@userId", userId);
        command.Parameters.Add("@info", SqlDbType.NVarChar, -1).Value = userInfo;
        await command.ExecuteNonQueryAsync(cancel);
    }
}
