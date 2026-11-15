using System;
using System.IO;

/// <summary>
/// The response filter of a copied request: what the response writes goes to SharePoint's stream first, then to the copy.
/// Only the copy is in a try: a failure of SharePoint's stream is SharePoint's, a failure of the copy is the module's
/// alone.
/// </summary>
sealed class TapFilter(Stream inner, TapCapture capture) : Stream
{
    static int noted;
    bool closed;
    int writing;

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        writing++;
        try
        {
            inner.Write(buffer, offset, count);
        }
        finally
        {
            writing--;
        }

        // a write inside a write: what this filter wrote made ASP.NET flush and filter it again (unbuffered at the final
        // filtering, as in IIS Express when a module turns buffering off after the handler), so the same bytes come
        // back before the outer write returns. Copied once, by the outermost; SharePoint's stream gets every write
        if (writing > 0)
        {
            if (Interlocked.Exchange(ref noted, 1) == 0)
            {
                TapLog.Write($"The response was filtered again inside a write of the module's filter (a note, once): copied once. {Environment.StackTrace}");
            }

            return;
        }

        try
        {
            capture.Writing(buffer, offset, count);
        }
        catch (Exception exception)
        {
            capture.Fail(exception);
        }
    }

    public override void Flush() => inner.Flush();

    public override void Close()
    {
        if (!closed)
        {
            closed = true;
            try
            {
                capture.Closing();
            }
            catch (Exception exception)
            {
                capture.Fail(exception);
            }
        }

        inner.Close();
        base.Close();
    }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();
}
