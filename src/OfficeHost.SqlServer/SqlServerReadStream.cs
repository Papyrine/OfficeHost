/// <summary>
/// The content of a file, read from the database in chunks, each by a statement on a connection of its own: a slow
/// reader holds no connection between chunks, nor a lock. A write of the file while it is read fails the read, rather
/// than mixing two versions.
/// </summary>
sealed class SqlServerReadStream(string connectionString, string table, string fileId, Guid version, long length) :
    Stream
{
    const int chunkSize = 4 * 1024 * 1024;

    long position;
    byte[] chunk = [];
    int chunkOffset;

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => length;

    public override long Position
    {
        get => position;
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        Read(buffer.AsSpan(offset, count));

    public override int Read(Span<byte> buffer)
    {
        if (!Available())
        {
            if (position >= length)
            {
                return 0;
            }

            using var connection = new SqlConnection(connectionString);
            connection.Open();
            using var command = ChunkCommand(connection);
            using var reader = command.ExecuteReader(CommandBehavior.SingleRow);
            Next(reader.Read(), reader);
        }

        return Copy(buffer);
    }

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, Cancel cancel) =>
        ReadAsync(buffer.AsMemory(offset, count), cancel).AsTask();

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, Cancel cancel = default)
    {
        if (!Available())
        {
            if (position >= length)
            {
                return 0;
            }

            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancel);
            await using var command = ChunkCommand(connection);
            await using var reader = await command.ExecuteReaderAsync(CommandBehavior.SingleRow, cancel);
            Next(await reader.ReadAsync(cancel), reader);
        }

        return Copy(buffer.Span);
    }

    bool Available() => chunkOffset < chunk.Length;

    SqlCommand ChunkCommand(SqlConnection connection)
    {
        // substring is 1-based
        var command = new SqlCommand($"select Version, substring(Content, @offset + 1, @count) from {table} where Id = @id", connection);
        command.Parameters.Add("@id", SqlDbType.NVarChar, 450).Value = fileId;
        command.Parameters.Add("@offset", SqlDbType.BigInt).Value = position;
        command.Parameters.Add("@count", SqlDbType.BigInt).Value = (long) chunkSize;
        return command;
    }

    void Next(bool found, SqlDataReader reader)
    {
        if (!found)
        {
            throw new IOException($"The file {fileId} was deleted while it was read");
        }

        if (reader.GetGuid(0) != version)
        {
            throw new IOException($"The file {fileId} was written while it was read");
        }

        chunk = (byte[]) reader[1];
        chunkOffset = 0;
        if (chunk.Length == 0)
        {
            throw new IOException($"The file {fileId} is shorter than when it was opened");
        }
    }

    int Copy(Span<byte> buffer)
    {
        var count = Math.Min(buffer.Length, chunk.Length - chunkOffset);
        chunk.AsSpan(chunkOffset, count).CopyTo(buffer);
        chunkOffset += count;
        position += count;
        return count;
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    public override void SetLength(long value) =>
        throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();
}
