/// <summary>
/// The content of a file, read from the database as it is read, which keeps the reader and its connection open until
/// disposed.
/// </summary>
sealed class SqlServerReadStream(SqlConnection connection, SqlDataReader reader, Stream inner) :
    Stream
{
    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;
    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) =>
        inner.Read(buffer, offset, count);

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, Cancel cancel) =>
        inner.ReadAsync(buffer, offset, count, cancel);

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, Cancel cancel = default) =>
        inner.ReadAsync(buffer, cancel);

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    public override void SetLength(long value) =>
        throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            inner.Dispose();
            reader.Dispose();
            connection.Dispose();
        }

        base.Dispose(disposing);
    }

    public override async ValueTask DisposeAsync()
    {
        await inner.DisposeAsync();
        await reader.DisposeAsync();
        await connection.DisposeAsync();
        await base.DisposeAsync();
    }
}
