/// <summary>
/// The request stream that SharePoint's handler reads, when it reads the body without buffering it, so that nothing can
/// read it again (WCF, as SharePoint 2016 serves cellstorage.svc): what the handler reads passes through as it is read,
/// synchronously or not, and is copied as it passes: its start, where the SOAP envelope names the document, and all of it
/// to the copy when there is one. The copy takes nothing until the start says whether the document is sampled
/// (<paramref name="sampled"/>, null while it names none): a body of a document that is not sampled is never copied
/// beyond its start. A start that fills without naming one is copied on, until <see cref="Settle"/> says. Only the
/// copying is in a try: a failure of the stream is SharePoint's, a failure of the copy only stops it.
/// </summary>
sealed class TapInputStream(Stream inner, TapBuffer? copy, Func<string, bool?>? sampled = null) : Stream
{
    const int startSize = 64 * 1024;
    Lock gate = new();
    byte[] start = new byte[startSize];
    int startLength;
    bool stopped;
    // the copy has the start, and takes what is read
    bool copying;
    // whether the document is sampled, once known
    bool? wanted;

    /// <summary>
    /// The bytes the handler has read.
    /// </summary>
    public long Passed { get; private set; }

    /// <summary>
    /// Whether the handler read to the end of the body: a read that asked for bytes got none.
    /// </summary>
    public bool Ended { get; private set; }

    /// <summary>
    /// Why the copy stopped before the end, or null.
    /// </summary>
    public string? Failure { get; private set; }

    /// <summary>
    /// The first 64 KB the handler has read.
    /// </summary>
    public byte[] Start()
    {
        lock (gate)
        {
            var bytes = new byte[startLength];
            Buffer.BlockCopy(start, 0, bytes, 0, startLength);
            return bytes;
        }
    }

    /// <summary>
    /// Copies no more: the copy is sent or released.
    /// </summary>
    public void Stop()
    {
        lock (gate)
        {
            stopped = true;
        }
    }

    /// <summary>
    /// Whether the body is copied: as its start said when it named the document, else as <paramref name="otherwise"/>
    /// says, from here on. The copy then gets the start it waited with.
    /// </summary>
    public bool Settle(bool otherwise)
    {
        lock (gate)
        {
            if (wanted is { } earlier)
            {
                return earlier;
            }

            wanted = otherwise;
            if (otherwise &&
                !copying &&
                !stopped)
            {
                try
                {
                    Begin();
                }
                catch (Exception exception)
                {
                    Failure ??= $"copying the request body failed: {exception.Message}";
                    stopped = true;
                }
            }

            return otherwise;
        }
    }

    // under the gate: what was read so far is all in the start
    void Begin()
    {
        copying = true;
        copy?.Write(start, 0, startLength);
    }

    public override bool CanRead => inner.CanRead;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override bool CanTimeout => inner.CanTimeout;

    public override int ReadTimeout
    {
        get => inner.ReadTimeout;
        set => inner.ReadTimeout = value;
    }

    public override long Length => inner.Length;

    public override long Position
    {
        get => inner.Position;
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        var read = inner.Read(buffer, offset, count);
        Copy(buffer, offset, read, count);
        return read;
    }

    // the inner stream's own asynchronous read, as the handler would have had it: not Stream's, a synchronous Read on
    // the thread pool
    public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
    {
        var read = new PendingRead(buffer, offset, count, state);
        AsyncCallback? done = null;
        if (callback != null)
        {
            done = _ =>
            {
                // before the handler's callback, which ends the read: it may come before BeginRead returns
                read.Inner = _;
                callback(read);
            };
        }

        read.Inner = inner.BeginRead(buffer, offset, count, done, state);
        return read;
    }

    public override int EndRead(IAsyncResult asyncResult)
    {
        var read = (PendingRead) asyncResult;
        var count = inner.EndRead(read.Inner!);
        Copy(read.Buffer, read.Offset, count, read.Count);
        return count;
    }

    /// <param name="count">The bytes read.</param>
    /// <param name="requested">The bytes the handler asked for.</param>
    void Copy(byte[] buffer, int offset, int count, int requested)
    {
        lock (gate)
        {
            if (count <= 0)
            {
                Ended |= requested > 0;
                return;
            }

            Passed += count;
            if (stopped)
            {
                return;
            }

            try
            {
                var kept = Math.Min(count, startSize - startLength);
                Buffer.BlockCopy(buffer, offset, start, startLength, kept);
                startLength += kept;
                if (copy == null ||
                    wanted == false)
                {
                    return;
                }

                if (copying)
                {
                    copy.Write(buffer, offset, count);
                    return;
                }

                // until here, all that was read is in the start
                wanted ??= sampled?.Invoke(Encoding.UTF8.GetString(start, 0, startLength));
                if (wanted == true ||
                    (wanted == null && startLength == startSize))
                {
                    Begin();
                    copy.Write(buffer, offset + kept, count - kept);
                }
            }
            catch (Exception exception)
            {
                Failure ??= $"copying the request body failed: {exception.Message}";
                stopped = true;
            }
        }
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// A read of <see cref="BeginRead"/>: where its bytes go, and the inner stream's read.
    /// </summary>
    sealed class PendingRead(byte[] buffer, int offset, int count, object? state) : IAsyncResult
    {
        public byte[] Buffer { get; } = buffer;

        public int Offset { get; } = offset;

        public int Count { get; } = count;

        public IAsyncResult? Inner { get; set; }

        public object? AsyncState => state;

        public WaitHandle AsyncWaitHandle => Inner!.AsyncWaitHandle;

        public bool CompletedSynchronously => Inner!.CompletedSynchronously;

        public bool IsCompleted => Inner?.IsCompleted ?? false;
    }
}
