/// <summary>
/// A stream read to at most a number of bytes: reading past it throws <see cref="InvalidDataException"/>.
/// </summary>
sealed class LimitedStream(Stream inner, long limit) :
    Stream
{
    long read;

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => read;
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
        read += count;
        if (read > limit)
        {
            throw new InvalidDataException($"More than {limit} bytes");
        }

        return count;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
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
