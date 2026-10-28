using System;
using System.IO;

/// <summary>
/// The response filter of a copied request: what the response writes goes to SharePoint's stream first, then to the copy.
/// Only the copy is in a try: a failure of SharePoint's stream is SharePoint's, a failure of the copy is the module's
/// alone.
/// </summary>
sealed class TapFilter(Stream inner, TapCapture capture) : Stream
{
    bool closed;

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
        inner.Write(buffer, offset, count);
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
