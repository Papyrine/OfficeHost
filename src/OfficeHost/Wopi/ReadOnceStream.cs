/// <summary>
/// A request body that tells whether its reading started: once it has, it cannot be given to a storage again.
/// </summary>
sealed class ReadOnceStream(Stream inner) :
    Stream
{
    public bool Started { get; private set; }

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
        Read(buffer.AsSpan(offset, count));

    public override int Read(Span<byte> buffer) =>
        Count(inner.Read(buffer));

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, Cancel cancel) =>
        ReadAsync(buffer.AsMemory(offset, count), cancel).AsTask();

    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, Cancel cancel = default) =>
        Count(await inner.ReadAsync(buffer, cancel));

    int Count(int count)
    {
        if (count > 0)
        {
            Started = true;
        }

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
