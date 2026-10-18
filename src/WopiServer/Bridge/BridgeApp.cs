/// <summary>
/// The HTTP clients used by the bridge: Python <c>requests</c> follows redirects except for HEAD requests,
/// and some calls explicitly disable them.
/// </summary>
record BridgeHttp(HttpClient Follow, HttpClient NoRedirect);

/// <summary>
/// A fatal failure when contacting a bridged app. The message, when not empty, is shown to the user.
/// </summary>
class AppFailure(string message = "", Exception? inner = null) :
    Exception(message, inner);

/// <summary>
/// An application integrated through the WOPI bridge.
/// </summary>
abstract class BridgeApp(BridgeHttp http, WopiClient wopi, ILogger log)
{
    protected BridgeHttp Http { get; } = http;
    protected WopiClient Wopi { get; } = wopi;
    protected ILogger Log { get; } = log;

    /// <summary>
    /// The application name, as registered by the first open request or by the configuration.
    /// </summary>
    public string AppName { get; set; } = "";

    /// <summary>
    /// The URL of the app as seen by the users.
    /// </summary>
    public string AppExtUrl { get; private set; } = "";

    /// <summary>
    /// The URL of the app as seen by this server, which differs in containerized deployments.
    /// </summary>
    protected string AppUrl { get; private set; } = "";

    protected string? ApiKey { get; private set; }

    public bool DisableZip { get; set; }

    public async Task Init(string appUrl, string? appIntUrl, string? apiKey)
    {
        AppExtUrl = appUrl;
        AppUrl = appIntUrl ?? appUrl;
        ApiKey = apiKey;
        await Probe();
    }

    /// <summary>
    /// Validates the connection to the app. Throws <see cref="AppFailure"/>.
    /// </summary>
    protected abstract Task Probe();

    /// <summary>
    /// Copies the document from storage to the app. A null <paramref name="docId"/> means a read-only open
    /// on a random docid. Returns the WOPI lock structure for the document.
    /// </summary>
    public abstract Task<BridgeLock> LoadFromStorage(JsonNode fileMetadata, string wopiSrc, string acctok, string? docId);

    /// <summary>
    /// A valid URL to the app for the given WOPI context.
    /// </summary>
    public abstract Task<string> GetRedirectUrl(ViewMode viewMode, string wopiSrc, string acctok, string docId, string filename, string displayName, string? revaToken);

    /// <summary>
    /// Fetches the document from the app. Throws <see cref="AppFailure"/>.
    /// </summary>
    public abstract Task<byte[]> Fetch(BridgeLock wopiLock, string acctok);

    /// <summary>
    /// Copies the document from the app back to storage.
    /// </summary>
    public abstract Task<BridgeReply> SaveToStorage(string wopiSrc, string acctok, bool isClose, BridgeLock wopiLock);

    protected static string Body(byte[] content, int length) =>
        Encoding.UTF8.GetString(content, 0, Math.Min(content.Length, length));

    protected static string Body(byte[] content) =>
        Encoding.UTF8.GetString(content);

    /// <summary>
    /// The last path segment of a redirect location.
    /// </summary>
    protected static string LocationId(HttpResponseMessage response)
    {
        var location = response.Headers.Location ?? throw new AppFailure();
        var path = location.IsAbsoluteUri ? location.AbsolutePath : location.OriginalString.Split('?', '#')[0];
        return path.Split('/')[^1];
    }

    protected static StringContent ContentOf(string value, string mediaType) =>
        new(value, Encoding.UTF8, mediaType);

    protected static string Query(params (string Name, string Value)[] parameters) =>
        UrlCodec.UrlEncode(parameters.Select(_ => KeyValuePair.Create(_.Name, _.Value)));

    protected static bool IsConnectionError(Exception exception) =>
        exception is HttpRequestException or TaskCanceledException;
}
