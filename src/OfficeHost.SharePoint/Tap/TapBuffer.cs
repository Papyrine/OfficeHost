/// <summary>
/// The bytes of copies the module holds, across all requests, bounded: a copy that would go beyond stops.
/// </summary>
sealed class TapBudget(long limit)
{
    long used;

    public long Used => Interlocked.Read(ref used);

    public bool TryTake(long bytes)
    {
        if (Interlocked.Add(ref used, bytes) <= limit)
        {
            return true;
        }

        Interlocked.Add(ref used, -bytes);
        return false;
    }

    public void Give(long bytes) => Interlocked.Add(ref used, -bytes);
}

/// <summary>
/// A copy of a body, in chunks of 64 KB (under the large object heap), each taken from the <see cref="TapBudget"/>. A
/// copy beyond its limit or the budget stops and gives its chunks back: <see cref="Overflowed"/>.
/// </summary>
sealed class TapBuffer(TapBudget budget, long limit)
{
    const int chunkSize = 64 * 1024;
    List<byte[]> chunks = [];
    int used = chunkSize;

    public long Length { get; private set; }

    public bool Overflowed { get; private set; }

    public void Write(byte[] buffer, int offset, int count)
    {
        while (count > 0 &&
               !Overflowed)
        {
            if (used == chunkSize)
            {
                if (Length >= limit ||
                    !budget.TryTake(chunkSize))
                {
                    Overflowed = true;
                    Release();
                    return;
                }

                chunks.Add(new byte[chunkSize]);
                used = 0;
            }

            var length = Math.Min(count, chunkSize - used);
            Buffer.BlockCopy(buffer, offset, chunks[^1], used, length);
            used += length;
            offset += length;
            count -= length;
            Length += length;
        }

        if (Length > limit)
        {
            Overflowed = true;
            Release();
        }
    }

    /// <summary>
    /// Copies a stream to the end.
    /// </summary>
    public void ReadFrom(Stream stream)
    {
        var buffer = new byte[chunkSize];
        int read;
        while (!Overflowed &&
               (read = stream.Read(buffer, 0, buffer.Length)) > 0)
        {
            Write(buffer, 0, read);
        }
    }

    public void CopyTo(Stream stream)
    {
        var left = Length;
        foreach (var chunk in chunks)
        {
            var length = (int) Math.Min(left, chunkSize);
            stream.Write(chunk, 0, length);
            left -= length;
        }
    }

    public byte[] ToArray()
    {
        using var stream = new MemoryStream();
        CopyTo(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Gives the chunks back to the budget: the copy is sent, dropped, or overflowed.
    /// </summary>
    public void Release()
    {
        budget.Give((long) chunks.Count * chunkSize);
        chunks.Clear();
        used = chunkSize;
        if (!Overflowed)
        {
            Length = 0;
        }
    }
}
