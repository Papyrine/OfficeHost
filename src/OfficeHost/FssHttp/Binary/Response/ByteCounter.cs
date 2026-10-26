/// <summary>
/// An output that counts the bytes written to it and keeps none: the size of FSSHTTPB without making it, in a buffer
/// that is reused.
/// </summary>
sealed class ByteCounter :
    IBufferWriter<byte>
{
    byte[] scratch = new byte[4096];

    public long Count { get; private set; }

    public void Reset() => Count = 0;

    public void Advance(int count) => Count += count;

    public Memory<byte> GetMemory(int sizeHint = 0)
    {
        Ensure(sizeHint);
        return scratch;
    }

    public Span<byte> GetSpan(int sizeHint = 0)
    {
        Ensure(sizeHint);
        return scratch;
    }

    void Ensure(int sizeHint)
    {
        if (sizeHint > scratch.Length)
        {
            scratch = new byte[sizeHint];
        }
    }
}
