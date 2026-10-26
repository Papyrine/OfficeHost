namespace OfficeHost;

/// <summary>
/// The WOPI locks in a SQL Server table, shared by all the instances of the application. Expiry is decided by the
/// clock of the application, as for <see cref="MemoryLockStore"/>. Lock ids are compared as they are, whatever the
/// collation of the database.
/// </summary>
public sealed class SqlServerLockStore(string connectionString, TimeProvider time, string schema = "dbo") :
    IOfficeLockStore
{
    string table = SqlServerTables.Name(schema, SqlServerTables.Locks);

    public async Task<OfficeLock?> Get(string fileId, Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand($"select LockId, Expires from {table} where FileId = @fileId and Expires > @now", connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@now", time.GetUtcNow());
        return await ReadLock(command, cancel);
    }

    /// <summary>
    /// The longest lock id, as the table holds it.
    /// </summary>
    const int maxLockLength = 1024;

    public async Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, Cancel cancel = default)
    {
        if (value?.Id.Length > maxLockLength)
        {
            throw new ArgumentException($"A lock id is at most {maxLockLength} characters", nameof(value));
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
             declare @id nvarchar(1024), @expires datetimeoffset;
             select @id = LockId, @expires = Expires from {table} with (updlock, holdlock) where FileId = @fileId;
             if @expires <= @now
             begin
                 delete from {table} where FileId = @fileId;
                 set @id = null;
                 set @expires = null;
             end;
             if @id = @comparand collate Latin1_General_100_BIN2 or (@id is null and @comparand is null)
             begin
                 delete from {table} where FileId = @fileId;
                 if @value is not null
                     insert into {table} (FileId, LockId, Expires) values (@fileId, @value, @valueExpires);
             end;
             commit;
             select @id, @expires where @id is not null;
             """,
            connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@now", time.GetUtcNow());
        command.Parameters.Add("@comparand", SqlDbType.NVarChar, 1024).Value = (object?) comparand ?? DBNull.Value;
        command.Parameters.Add("@value", SqlDbType.NVarChar, 1024).Value = (object?) value?.Id ?? DBNull.Value;
        command.Parameters.Add("@valueExpires", SqlDbType.DateTimeOffset).Value = (object?) value?.Expires ?? DBNull.Value;
        return await ReadLock(command, cancel);
    }

    static async Task<OfficeLock?> ReadLock(SqlCommand command, Cancel cancel)
    {
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            return null;
        }

        return new(reader.GetString(0), reader.GetDateTimeOffset(1));
    }
}
