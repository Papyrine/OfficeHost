using System;

/// <summary>
/// What the module sends of a request: an exchange, SharePoint's request and response as raw HTTP as a recording has them
/// (docs/replay.md), or a tombstone, the document of an exchange it could not send and why (docs/shadow.md, "Protocol").
/// </summary>
sealed class TapItem
{
    TapItem(long sequence, DateTime started)
    {
        Sequence = sequence;
        Started = started;
    }

    public long Sequence { get; }

    public DateTime Started { get; }

    public byte[] RequestHead { get; private set; } = [];

    public TapBuffer? RequestBody { get; private set; }

    public byte[] ResponseHead { get; private set; } = [];

    public TapBuffer? ResponseBody { get; private set; }

    public string? Document { get; private set; }

    public string? Reason { get; private set; }

    public bool IsTombstone => Reason != null;

    public long RequestLength => RequestHead.Length + (RequestBody?.Length ?? 0);

    public long Length => RequestLength + ResponseHead.Length + (ResponseBody?.Length ?? 0);

    /// <param name="document">The URL of the document the request names, for a tombstone of it.</param>
    public static TapItem Exchange(long sequence, DateTime started, string? document, byte[] requestHead, TapBuffer? requestBody, byte[] responseHead, TapBuffer? responseBody) =>
        new(sequence, started)
        {
            Document = document,
            RequestHead = requestHead,
            RequestBody = requestBody,
            ResponseHead = responseHead,
            ResponseBody = responseBody
        };

    public static TapItem Tombstone(long sequence, DateTime started, string? document, string reason) =>
        new(sequence, started)
        {
            Document = document,
            Reason = reason
        };

    /// <summary>
    /// The tombstone of this exchange, its copies given back: it could not be sent.
    /// </summary>
    public TapItem Bury(string reason)
    {
        Release();
        return Tombstone(Sequence, Started, Document, reason);
    }

    public void Release()
    {
        RequestBody?.Release();
        ResponseBody?.Release();
    }
}
