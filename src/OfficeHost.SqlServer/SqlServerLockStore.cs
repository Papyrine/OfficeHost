namespace OfficeHost;

/// <summary>
/// The WOPI locks in a SQL Server table, shared by all the instances of the application. Expiry is decided by the
/// clock of the application, as for <see cref="MemoryLockStore"/>. Lock ids are compared as they are, whatever the
/// collation of the database. The holder of a lock is kept in the HolderId, HolderName and HolderClient columns, for
/// the application to read.
/// </summary>
public sealed class SqlServerLockStore(string connectionString, TimeProvider time, string schema = "dbo") :
    IOfficeLockStore
{
    string table = SqlServerTables.Name(schema, SqlServerTables.Locks);

    const string columns = "LockId, Expires, HolderId, HolderName, HolderClient";

    public Task<OfficeLock?> Get(string fileId, Cancel cancel = default) =>
        DeadlockRetry.Once(() => GetAttempt(fileId, cancel));

    async Task<OfficeLock?> GetAttempt(string fileId, Cancel cancel)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand(
            $"select {columns} from {table} where FileId = @fileId and Expires > @now",
            connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@now", time.GetUtcNow());
        return await ReadLock(command, cancel);
    }

    /// <summary>
    /// The longest lock id, as the table holds it.
    /// </summary>
    const int maxLockLength = 1024;

    /// <summary>
    /// The longest user id, as the table holds it.
    /// </summary>
    const int maxIdLength = 450;

    public Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, Cancel cancel = default) =>
        CompareExchange(fileId, comparand, value, null, cancel);

    public Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, OfficeUser? user, Cancel cancel = default) =>
        DeadlockRetry.Once(() => CompareExchangeAttempt(fileId, comparand, value, user ?? value?.Holder, cancel));

    async Task<OfficeLock?> CompareExchangeAttempt(string fileId, string? comparand, OfficeLock? value, OfficeUser? holder, Cancel cancel)
    {
        if (value?.Id.Length > maxLockLength)
        {
            throw new ArgumentException($"A lock id is at most {maxLockLength} characters", nameof(value));
        }

        if (holder?.Id.Length > maxIdLength)
        {
            throw new ArgumentException($"A user id is at most {maxIdLength} characters", "user");
        }

        // a longer comparand is no lock the table holds, and the size of the parameter would cut it into one
        if (comparand?.Length > maxLockLength)
        {
            return await Get(fileId, cancel);
        }

        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        // the range lock of the read holds until the commit, so racers on the same file go one at a time
        await using var command = new SqlCommand(
            $"""
             set xact_abort on;
             begin transaction;
             declare @id nvarchar(1024), @expires datetimeoffset, @holderId nvarchar(450), @holderName nvarchar(max), @holderClient nvarchar(16);
             select @id = LockId, @expires = Expires, @holderId = HolderId, @holderName = HolderName, @holderClient = HolderClient from {table} with (updlock, holdlock) where FileId = @fileId;
             if @expires <= @now
             begin
                 delete from {table} where FileId = @fileId;
                 set @id = null;
                 set @expires = null;
                 set @holderId = null;
                 set @holderName = null;
                 set @holderClient = null;
             end;
             if @id = @comparand collate Latin1_General_100_BIN2 or (@id is null and @comparand is null)
             begin
                 delete from {table} where FileId = @fileId;
                 if @value is not null
                     insert into {table} (FileId, LockId, Expires, HolderId, HolderName, HolderClient) values (@fileId, @value, @valueExpires, @valueHolderId, @valueHolderName, @valueHolderClient);
             end;
             commit;
             select @id, @expires, @holderId, @holderName, @holderClient where @id is not null;
             """,
            connection);
        var parameters = command.Parameters;
        parameters.AddWithValue("@fileId", fileId);
        parameters.AddWithValue("@now", time.GetUtcNow());
        parameters.Add("@comparand", SqlDbType.NVarChar, maxLockLength).Value = (object?) comparand ?? DBNull.Value;
        parameters.Add("@value", SqlDbType.NVarChar, maxLockLength).Value = (object?) value?.Id ?? DBNull.Value;
        parameters.Add("@valueExpires", SqlDbType.DateTimeOffset).Value = (object?) value?.Expires ?? DBNull.Value;
        parameters.Add("@valueHolderId", SqlDbType.NVarChar, maxIdLength).Value = (object?) holder?.Id ?? DBNull.Value;
        parameters.Add("@valueHolderName", SqlDbType.NVarChar, -1).Value = (object?) holder?.Name ?? DBNull.Value;
        parameters.Add("@valueHolderClient", SqlDbType.NVarChar, 16).Value = (object?) holder?.Client.ToString() ?? DBNull.Value;
        return await ReadLock(command, cancel);
    }

    static async Task<OfficeLock?> ReadLock(SqlCommand command, Cancel cancel)
    {
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            return null;
        }

        return new(reader.GetString(0), reader.GetDateTimeOffset(1))
        {
            Holder = ReadHolder(reader)
        };
    }

    /// <summary>
    /// The holder of the row, or null for a lock written before the holder was kept, or with no user.
    /// </summary>
    static OfficeUser? ReadHolder(SqlDataReader reader)
    {
        if (reader.IsDBNull(2))
        {
            return null;
        }

        string? name = null;
        if (!reader.IsDBNull(3))
        {
            name = reader.GetString(3);
        }

        var client = OfficeClient.Wopi;
        if (!reader.IsDBNull(4) &&
            Enum.TryParse<OfficeClient>(reader.GetString(4), out var parsed))
        {
            client = parsed;
        }

        return new(reader.GetString(2), name, client);
    }
}
