/// <summary>
/// What the module in SharePoint (OfficeHost.SharePoint) and the shadow send each other (docs/shadow.md, "Protocol").
/// A batch is a POST to <c>{pattern}/exchanges</c> of a <c>multipart/mixed</c> body, optionally gzip, of exchanges and
/// tombstones; a batch of none is a heartbeat. The shadow answers 202, or 503 with Retry-After when it is full or
/// paused, with <c>X-Shadow-Sample</c> when it sets the sampling.
/// </summary>
static class ShadowProtocol
{
    /// <summary>
    /// The key of <see cref="OfficeHostShadowOptions.ApiKey"/>.
    /// </summary>
    public const string Key = "X-OfficeHost-Shadow-Key";

    /// <summary>
    /// The batch: the instance of the module (a GUID per app domain), its server, the watermark (all exchanges it
    /// started before it are in this batch or before, or tombstoned), the exchanges it dropped without a tombstone since
    /// its last batch, and whether it is the last batch of the instance (its app domain stops).
    /// </summary>
    public const string Instance = "X-Shadow-Instance";

    public const string Server = "X-Shadow-Server";

    public const string Watermark = "X-Shadow-Watermark";

    public const string Dropped = "X-Shadow-Dropped";

    public const string Final = "X-Shadow-Final";

    /// <summary>
    /// A part: its sequence in the instance, when SharePoint received the request, and its kind. An exchange's part is
    /// the request then the response, raw HTTP as a recording has them (docs/replay.md), the request
    /// <see cref="RequestLength"/> bytes long. A tombstone's part names its document and why the exchange is missing.
    /// </summary>
    public const string Sequence = "X-Shadow-Sequence";

    public const string Started = "X-Shadow-Started";

    public const string Kind = "X-Shadow-Kind";

    public const string RequestLength = "X-Shadow-Request-Length";

    public const string Document = "X-Shadow-Document";

    public const string Reason = "X-Shadow-Reason";

    /// <summary>
    /// The answer: the percentage of documents the module sends.
    /// </summary>
    public const string Sample = "X-Shadow-Sample";
}
