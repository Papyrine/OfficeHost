/// <summary>
/// The request stream that SharePoint's handler reads, when it reads the body without buffering it, so that nothing can
/// read it again (WCF, as SharePoint 2016 serves cellstorage.svc): what the handler reads passes through as it is read,
/// synchronously or not, and is copied as it passes: its start, where the SOAP envelope names the document, and all of it
/// to the copy when there is one. Only the copying is in a try: a failure of the stream is SharePoint's, a failure of the
/// copy only stops it.
/// </summary>
sealed class TapInputStream(Stream inner, TapBuffer? copy) : Stream
{
    const int startSize = 64 * 1024;
    Lock gate = new();
    byte[] start = new byte[startSize];
    int startLength;
    bool stopped;

    /// <summary>
    /// The bytes the handler has read.
    /// </summary>
    public long Passed { get; private set; }

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
        Copy(buffer, offset, read);
        return read;
    }

    // the inner stream's own asynchronous read, as the handler would have had it: not Stream's, a synchronous Read on
    // the thread pool
    public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback? callback, object? state)
    {
        var read = new PendingRead(buffer, offset, state);
        AsyncCallback? done = null;
        if (callback != null)
        {
            done = result =>
            {
                // before the handler's callback, which ends the read: it may come before BeginRead returns
                read.Inner = result;
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
        Copy(read.Buffer, read.Offset, count);
        return count;
    }

    void Copy(byte[] buffer, int offset, int count)
    {
        if (count <= 0)
        {
            return;
        }

        lock (gate)
        {
            Passed += count;
            if (stopped)
            {
                return;
            }

            try
            {
                if (startLength < startSize)
                {
                    var length = Math.Min(count, startSize - startLength);
                    Buffer.BlockCopy(buffer, offset, start, startLength, length);
                    startLength += length;
                }

                copy?.Write(buffer, offset, count);
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
    sealed class PendingRead(byte[] buffer, int offset, object? state) : IAsyncResult
    {
        public byte[] Buffer { get; } = buffer;

        public int Offset { get; } = offset;

        public IAsyncResult? Inner { get; set; }

        public object? AsyncState => state;

        public WaitHandle AsyncWaitHandle => Inner!.AsyncWaitHandle;

        public bool CompletedSynchronously => Inner!.CompletedSynchronously;

        public bool IsCompleted => Inner?.IsCompleted ?? false;
    }
}
