/// <summary>
/// The settings of the WOPI endpoints, given to <see cref="WopiExtensions.AddWopi"/>.
/// </summary>
public class WopiOptions
{
    /// <summary>
    /// The name of the <see cref="HttpClient"/> that fetches the discovery of the WOPI clients, to configure a proxy or
    /// a handler with <c>services.AddHttpClient(WopiOptions.HttpClientName)</c>.
    /// </summary>
    public const string HttpClientName = "Wopi";

    /// <summary>
    /// The key that signs the access tokens of the built-in <see cref="IWopiTokens"/>, at least 32 characters. Not used
    /// when another <see cref="IWopiTokens"/> is registered.
    /// </summary>
    public string? SigningKey { get; set; }

    /// <summary>
    /// How long an access token is valid. Microsoft recommends 10 hours.
    /// </summary>
    public TimeSpan TokenLifetime { get; set; } = TimeSpan.FromHours(10);

    /// <summary>
    /// How long a lock lasts unless refreshed: [MS-WOPI] 3.3.5.1.13 sets it to 30 minutes.
    /// </summary>
    public TimeSpan LockDuration { get; set; } = TimeSpan.FromMinutes(30);

    /// <summary>
    /// The URL the WOPI clients reach this application at: the part of WOPISrc before <c>/wopi/files/</c>, and the base
    /// of the URL the proof keys sign. Needed when the WOPI clients reach the application by another URL than the
    /// requests do, as behind a reverse proxy, or when launching outside of a request. Defaults to the scheme, host and
    /// path base of the current request.
    /// </summary>
    public string? BaseUrl { get; set; }

    /// <summary>
    /// The discovery URLs of the trusted WOPI clients, as <c>https://collabora:9980/hosting/discovery</c>. Their actions
    /// are what <see cref="WopiLauncher"/> launches, and their proof keys validate the X-WOPI-Proof of the requests.
    /// </summary>
    public IList<string> DiscoveryUrls { get; } = [];

    /// <summary>
    /// Refuses the requests without X-WOPI-Proof. Only for WOPI clients that sign their requests, which Collabora
    /// Online does not.
    /// </summary>
    public bool RequireProof { get; set; }

    /// <summary>
    /// The largest file PutFile and PutRelativeFile accept, in bytes, or null for no limit. Office documents can be
    /// large, so this replaces the request body limit of the server for those requests.
    /// </summary>
    public long? MaxFileSize { get; set; }

    /// <summary>
    /// The X-WOPI-ServerVersion of the responses. Defaults to the version of this library.
    /// </summary>
    public string? ServerVersion { get; set; }

    /// <summary>
    /// The HostEditUrl of a file: the page of the application that opens the file in the editor.
    /// </summary>
    public Func<WopiFile, WopiAccess, string?>? HostEditUrl { get; set; }

    /// <summary>
    /// The HostViewUrl of a file: the page of the application that opens the file in the viewer.
    /// </summary>
    public Func<WopiFile, WopiAccess, string?>? HostViewUrl { get; set; }

    /// <summary>
    /// Adds or changes the properties of a CheckFileInfo response, such as CloseUrl, the breadcrumb or the branding.
    /// </summary>
    public Func<CheckFileInfoContext, Task>? OnCheckFileInfo { get; set; }
}

/// <summary>
/// The CheckFileInfo response being built, given to <see cref="WopiOptions.OnCheckFileInfo"/>.
/// </summary>
public class CheckFileInfoContext(HttpContext httpContext, WopiFile file, WopiAccess access, JsonObject properties)
{
    public HttpContext HttpContext { get; } = httpContext;

    public WopiFile File { get; } = file;

    public WopiAccess Access { get; } = access;

    /// <summary>
    /// The properties of the response, as [MS-WOPI] 3.3.5.1.2.2 names them.
    /// </summary>
    public JsonObject Properties { get; } = properties;
}
