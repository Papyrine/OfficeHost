using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web;

/// <summary>
/// The copy of a request and its response: the headers allowed, the user and the time when the request is authenticated;
/// at the response's first write (SharePoint's handler has read the request by then), whether its document is sampled,
/// and if so the request body, read again from its start; the response as it is written; its status and head once it is
/// complete; and the exchange, or a tombstone, to the queue once the request has completed. The body of a shared access
/// request is copied when the request is authenticated instead: after its ASMX handler, it reads empty. While the shadow
/// pauses the module or cannot be reached, a request of a sampled document is not copied but becomes a tombstone, so that
/// the shadow knows its document misses an exchange.
/// </summary>
sealed class TapCapture
{
    // the headers of the request a replay needs; never credentials, cookies or the request digest. Host, as the shadow
    // replays only the documents of the SharePoint the request went to
    static string[] requestHeaders =
    [
        "Host", "Content-Type", "SOAPAction", "User-Agent", "Scenario", "ScenarioType", "X-RequestStats", "Application",
        "FileExtension", "X-Office-Major-Version"
    ];

    const int peek = 64 * 1024;
    HttpContext context;
    TapRuntime runtime;
    string kind;
    bool copying;
    bool unread;
    long sequence;
    DateTime started;
    byte[] requestHead;
    int decided;
    string? document;
    TapBuffer? response;
    TapBuffer? requestBody;
    byte[]? responseHead;
    int status;
    bool closed;
    string? failure;

    /// <param name="copying">Whether to copy, else to send a tombstone of a sampled document.</param>
    public TapCapture(HttpContext context, TapRuntime runtime, string kind, bool copying)
    {
        this.context = context;
        this.runtime = runtime;
        this.kind = kind;
        this.copying = copying;
        sequence = runtime.NextSequence();
        started = TapClock.UtcNow;
        requestHead = RequestHead(context, started);
        if (kind == "sharedaccess")
        {
            // read before SharePoint, through the buffered input stream as the ASMX handler reads it: IsOnlyClient is a
            // few hundred bytes, and the lab's recorder read it so with no change to SharePoint's answers
            CopyRequestBody();
        }

        runtime.Queue.Started(sequence, started);
    }

    /// <summary>
    /// The service of a path: <c>cellstorage</c>, <c>sharedaccess</c>, or null for one not copied.
    /// </summary>
    public static string? KindOf(string path)
    {
        if (path.EndsWith("/_vti_bin/cellstorage.svc/CellStorageService", StringComparison.OrdinalIgnoreCase) ||
            path.EndsWith("/_vti_bin/cobalt.ashx/cellstorageservice", StringComparison.OrdinalIgnoreCase))
        {
            return "cellstorage";
        }

        if (path.EndsWith("/_vti_bin/sharedaccess.asmx", StringComparison.OrdinalIgnoreCase))
        {
            return "sharedaccess";
        }

        return null;
    }

    static byte[] RequestHead(HttpContext context, DateTime started)
    {
        var request = context.Request;
        var head = new StringBuilder();
        head.Append(request.HttpMethod).Append(' ').Append(request.RawUrl).Append(" HTTP/1.1\r\n");
        foreach (var name in requestHeaders)
        {
            if (request.Headers[name] is { } value)
            {
                head.Append(name).Append(": ").Append(value).Append("\r\n");
            }
        }

        head.Append("X-Replay-User: ").Append(context.User?.Identity?.Name ?? "").Append("\r\n");
        head.Append("X-Replay-Time: ").Append(started.ToString("o", CultureInfo.InvariantCulture)).Append("\r\n\r\n");
        return Encoding.UTF8.GetBytes(head.ToString());
    }

    /// <summary>
    /// What the response writes, after it went to SharePoint's stream.
    /// </summary>
    public void Writing(byte[] buffer, int offset, int count)
    {
        if (decided == 0)
        {
            Decide();
        }

        if (decided == 1)
        {
            response?.Write(buffer, offset, count);
        }
    }

    /// <summary>
    /// Whether the document is sampled: IsOnlyClient always (it names no document, and is small), a cell storage
    /// request by its document's URL. Not sampled, it is not copied at all. Decided once, even when the body cannot be
    /// read: then the exchange is lost to a document unknown.
    /// </summary>
    void Decide()
    {
        var sampled = true;
        if (kind == "cellstorage")
        {
            try
            {
                document = TapSampler.Document(Peek());
            }
            catch (Exception exception)
            {
                decided = 1;
                unread = true;
                failure = $"the request body could not be read again: {exception.Message}";
                return;
            }

            sampled = TapSampler.Sampled(TapSampler.Key(document ?? ""), runtime.Sender.SamplePercent);
        }

        if (!sampled)
        {
            decided = 2;
            runtime.Queue.Ended(sequence);
            return;
        }

        decided = 1;
        if (!copying)
        {
            failure = "not copied: the shadow paused the module, or could not be reached";
            return;
        }

        response = new(runtime.Budget, runtime.Settings.MaxExchangeBytes);
        if (requestBody == null)
        {
            CopyRequestBody();
        }
    }

    /// <summary>
    /// The request body, read again from its start: at the response's first write, SharePoint's handler (cobalt.ashx
    /// for cellstorage.svc) has read it, and reading it before made PowerPoint drop a request.
    /// </summary>
    void CopyRequestBody()
    {
        requestBody = new(runtime.Budget, runtime.Settings.MaxExchangeBytes);
        var input = context.Request.InputStream;
        var position = input.Position;
        try
        {
            input.Position = 0;
            requestBody.ReadFrom(input);
        }
        catch (Exception exception) when (exception is HttpException or IOException or InvalidOperationException or NotSupportedException)
        {
            // a handler that read the body without buffering it (SharePoint 2016 or 2019?)
            failure = $"the request body could not be read again: {exception.Message}";
        }
        finally
        {
            input.Position = position;
        }
    }

    /// <summary>
    /// The start of the request body, where the SOAP envelope names the document: read again from its start, as
    /// SharePoint's handler has read it.
    /// </summary>
    string Peek()
    {
        var input = context.Request.InputStream;
        var position = input.Position;
        try
        {
            input.Position = 0;
            var buffer = new byte[peek];
            var length = 0;
            int read;
            while (length < buffer.Length &&
                   (read = input.Read(buffer, length, buffer.Length - length)) > 0)
            {
                length += read;
            }

            return Encoding.UTF8.GetString(buffer, 0, length);
        }
        finally
        {
            input.Position = position;
        }
    }

    /// <summary>
    /// The response is complete (ASP.NET closes the filter): its status and head.
    /// </summary>
    public void Closing()
    {
        if (decided == 0)
        {
            Decide();
        }

        closed = true;
        if (decided != 1)
        {
            return;
        }

        var httpResponse = context.Response;
        status = httpResponse.StatusCode;
        if (status == 401)
        {
            return;
        }

        responseHead = ResponseHead(httpResponse);
    }

    static byte[] ResponseHead(HttpResponse response)
    {
        var head = new StringBuilder();
        head.Append("HTTP/1.1 ").Append(response.StatusCode).Append(' ').Append(response.StatusDescription).Append("\r\n");
        if (response.ContentType is { Length: > 0 } type)
        {
            if (response.Charset is { Length: > 0 } charset &&
                !type.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase) &&
                type.IndexOf("charset", StringComparison.OrdinalIgnoreCase) < 0)
            {
                type += "; charset=" + charset;
            }

            head.Append("Content-Type: ").Append(type).Append("\r\n");
        }

        head.Append("\r\n");
        return Encoding.UTF8.GetBytes(head.ToString());
    }

    /// <summary>
    /// Copying failed: nothing more is copied, and the exchange becomes a tombstone.
    /// </summary>
    public void Fail(Exception exception)
    {
        failure ??= $"copying failed: {exception.Message}";
        TapLog.Write($"Copying failed: {exception}");
    }

    /// <summary>
    /// The request has completed: the exchange to the queue, or a tombstone when it could not be copied in full.
    /// </summary>
    public void Completed()
    {
        try
        {
            if (Enqueue())
            {
                runtime.Sender.Signal();
            }
        }
        catch (Exception exception)
        {
            Release();
            TapLog.Write($"A request was not copied: {exception.Message}");
        }
        finally
        {
            // only once its exchange or tombstone waits: the watermark must not pass its start before
            runtime.Queue.Ended(sequence);
        }
    }

    /// <summary>
    /// The exchange, or its tombstone, to the queue; false when there is nothing to send.
    /// </summary>
    bool Enqueue()
    {
        if (decided != 1 ||
            status == 401)
        {
            Release();
            return false;
        }

        var reason = Reason();
        if (reason == null)
        {
            var item = TapItem.Exchange(sequence, started, document, requestHead, requestBody, responseHead!, response);
            if (!runtime.Queue.TryAdd(item))
            {
                runtime.Queue.AddTombstone(item.Bury("the module's queue was full"));
            }

            return true;
        }

        Release();
        if (unread)
        {
            // no document to name: every document of this instance is untrusted
            runtime.Queue.Drop();
            return true;
        }

        runtime.Queue.AddTombstone(TapItem.Tombstone(sequence, started, document, reason));
        return true;
    }

    string? Reason()
    {
        if (!closed)
        {
            return "the response did not complete";
        }

        if (failure != null)
        {
            return failure;
        }

        if (response!.Overflowed ||
            requestBody!.Overflowed ||
            response.Length + requestBody.Length > runtime.Settings.MaxExchangeBytes)
        {
            return "larger than the module copies, or its budget full";
        }

        if (kind == "cellstorage" &&
            requestBody.Length == 0)
        {
            return "no request body";
        }

        return null;
    }

    void Release()
    {
        response?.Release();
        requestBody?.Release();
    }
}
