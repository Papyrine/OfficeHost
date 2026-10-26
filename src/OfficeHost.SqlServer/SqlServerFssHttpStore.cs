namespace OfficeHost;

/// <summary>
/// The states of desktop Office coauthoring in a SQL Server table, shared by all the instances of the application.
/// </summary>
public sealed class SqlServerFssHttpStore(string connectionString, string schema = "dbo") :
    IFssHttpStore
{
    string table = SqlServerTables.Name(schema, SqlServerTables.FssHttpStates);

    public async Task<FssHttpState?> Get(string fileId, Guid partition, Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand($"select Version, Data from {table} where FileId = @fileId and Partition = @partition", connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@partition", partition);
        await using var reader = await command.ExecuteReaderAsync(cancel);
        if (!await reader.ReadAsync(cancel))
        {
            return null;
        }

        return new(reader.GetString(0), (byte[]) reader[1]);
    }

    public async Task<string?> GetVersion(string fileId, Guid partition, Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand($"select Version from {table} where FileId = @fileId and Partition = @partition", connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@partition", partition);
        return (string?) await command.ExecuteScalarAsync(cancel);
    }

    public async Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FssHttpState value, Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        var sql = $"update {table} set Version = @version, Data = @data where FileId = @fileId and Partition = @partition and Version = @comparand";
        if (comparand == null)
        {
            sql = $"""
                   insert into {table} (FileId, Partition, Version, Data)
                   select @fileId, @partition, @version, @data
                   where not exists (select 1 from {table} with (updlock, holdlock) where FileId = @fileId and Partition = @partition);
                   """;
        }

        await using var command = new SqlCommand(sql, connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        command.Parameters.AddWithValue("@partition", partition);
        command.Parameters.Add("@version", SqlDbType.NVarChar, -1).Value = value.Version;
        // the state as it is, without a copy: that of the content is the size of the file
        command.Parameters.Add("@data", SqlDbType.VarBinary, -1).Value = Data(value.Data);
        command.CommandTimeout = 0;
        command.Parameters.Add("@comparand", SqlDbType.NVarChar, -1).Value = (object?) comparand ?? DBNull.Value;
        try
        {
            return await command.ExecuteNonQueryAsync(cancel) == 1;
        }
        catch (SqlException exception) when (exception.IsDuplicateKey())
        {
            // inserted meanwhile
            return false;
        }
    }

    static object Data(ReadOnlyMemory<byte> data)
    {
        if (MemoryMarshal.TryGetArray(data, out var segment))
        {
            if (segment.Offset == 0 &&
                segment.Count == segment.Array!.Length)
            {
                return segment.Array;
            }

            return new MemoryStream(segment.Array!, segment.Offset, segment.Count, false);
        }

        return data.ToArray();
    }

    public async Task Delete(string fileId, Cancel cancel = default)
    {
        await using var connection = new SqlConnection(connectionString);
        await connection.OpenAsync(cancel);
        await using var command = new SqlCommand($"delete from {table} where FileId = @fileId", connection);
        command.Parameters.AddWithValue("@fileId", fileId);
        await command.ExecuteNonQueryAsync(cancel);
    }
}
