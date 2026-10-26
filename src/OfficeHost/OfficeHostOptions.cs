namespace OfficeHost;

/// <summary>
/// The settings of the WOPI endpoints, given to <see cref="OfficeHostExtensions.AddOfficeHost"/>.
/// </summary>
public class OfficeHostOptions
{
    /// <summary>
    /// The name of the <see cref="HttpClient"/> that fetches the discovery of the WOPI clients, to configure a proxy or
    /// a handler with <c>services.AddHttpClient(OfficeHostOptions.HttpClientName)</c>.
    /// </summary>
    public const string HttpClientName = "OfficeHost";

    /// <summary>
    /// The key that signs the access tokens of the built-in <see cref="IOfficeTokens"/>, at least 32 characters. Not used
    /// when another <see cref="IOfficeTokens"/> is registered.
    /// </summary>
    public string? SigningKey { get; set; }

    /// <summary>
    /// The keys that signed the tokens before <see cref="SigningKey"/>, to rotate it: tokens they signed stay valid until
    /// they expire, after <see cref="TokenLifetime"/>, when the key can be removed. New tokens are signed with
    /// <see cref="SigningKey"/>. To revoke tokens before they expire, register an <see cref="IOfficeTokens"/> that keeps
    /// them.
    /// </summary>
    public IList<string> PreviousSigningKeys { get; } = [];

    /// <summary>
    /// How long an access token is valid. Microsoft recommends 10 hours.
    /// </summary>
    public TimeSpan TokenLifetime { get; set; } = TimeSpan.FromHours(10);

    /// <summary>
    /// How long a lock lasts unless refreshed: [MS-WOPI] 3.3.5.1.13 sets it to 30 minutes.
    /// </summary>
    public TimeSpan LockDuration { get; set; } = TimeSpan.FromMinutes(30);

    /// <summary>
    /// The path the sites of desktop Office are under (<see cref="OfficeHostExtensions.MapFssHttp"/>): the site of a file
    /// is this path followed by a token, as <see cref="OfficeLauncher.LaunchDesktop"/> builds it.
    /// </summary>
    public string DesktopPath { get; set; } = "/fsshttp";

    /// <summary>
    /// The scheme, host and path base desktop Office reaches the application by, which runs on the machine of the user
    /// as the browser does: needed behind a reverse proxy, or when launching outside of a request. Defaults to those of
    /// the current request, not to <see cref="BaseUrl"/>, which is for WOPI clients.
    /// </summary>
    public string? DesktopBaseUrl { get; set; }

    /// <summary>
    /// The URL the WOPI clients reach this application at: the part of WOPISrc before <c>/wopi/files/</c>, and the base
    /// of the URL the proof keys sign. Needed when the WOPI clients reach the application by another URL than the
    /// requests do, as behind a reverse proxy, or when launching outside of a request. Defaults to the scheme, host and
    /// path base of the current request.
    /// </summary>
    public string? BaseUrl { get; set; }

    /// <summary>
    /// The discovery URLs of the trusted WOPI clients, as <c>https://collabora:9980/hosting/discovery</c>. Their actions
    /// are what <see cref="OfficeLauncher"/> launches, and their proof keys validate the X-WOPI-Proof of the requests.
    /// </summary>
    public IList<string> DiscoveryUrls { get; } = [];

    /// <summary>
    /// Refuses the requests without X-WOPI-Proof. Only for WOPI clients that sign their requests, which Collabora
    /// Online does not. Without it, an access token is enough to call the WOPI endpoints, so that the user of a
    /// <see cref="OfficeAccess.ViewOnly"/> token who takes it from the page can still download the file.
    /// </summary>
    public bool RequireProof { get; set; }

    /// <summary>
    /// The largest file PutFile and PutRelativeFile accept, in bytes, or null for no limit; 512 MB by default. Office
    /// documents can be large, so this replaces the request body limit of the server for those requests, and for the
    /// requests of desktop Office, which carry the changes of a file.
    /// </summary>
    public long? MaxFileSize { get; set; } = 512L * 1024 * 1024;

    /// <summary>
    /// The X-WOPI-ServerVersion of the responses. Defaults to the version of this library.
    /// </summary>
    public string? ServerVersion { get; set; }

    /// <summary>
    /// The HostEditUrl of a file: the page of the application that opens the file in the editor.
    /// </summary>
    public Func<OfficeFile, OfficeAccess, string?>? HostEditUrl { get; set; }

    /// <summary>
    /// The HostViewUrl of a file: the page of the application that opens the file in the viewer.
    /// </summary>
    public Func<OfficeFile, OfficeAccess, string?>? HostViewUrl { get; set; }

    /// <summary>
    /// Adds or changes the properties of a CheckFileInfo response, such as CloseUrl, the breadcrumb or the branding.
    /// </summary>
    public Func<CheckFileInfoContext, Task>? OnCheckFileInfo { get; set; }

    /// <summary>
    /// The value of a placeholder of the URL of an action, by its name in the discovery ([MS-WOPI] 3.1.5.1.1.2.3.1), as
    /// <c>UI_LLCC</c> (the language of the user interface), <c>DC_LLCC</c> (that of the data), <c>BUSINESS_USER</c> (1
    /// for a user of a business: Microsoft requires it), <c>THEME_ID</c> or <c>DISABLE_CHAT</c>; null drops it, as all are
    /// dropped without this.
    /// </summary>
    public Func<OfficeAccess, string, string?>? UrlPlaceholder { get; set; }

    /// <summary>
    /// The URL to share a file by, for GetShareUrl ([MS-WOPI] 3.3.5.1.8): given the file, the access of the user who
    /// shares it and the type of URL, <c>ReadOnly</c> or <c>ReadWrite</c>; null when the user cannot share it so. Without
    /// it, WOPI clients offer no sharing.
    /// </summary>
    public Func<OfficeFile, OfficeAccess, string, Task<string?>>? ShareUrl { get; set; }

    /// <summary>
    /// Called after a file was saved by a WOPI client (PutFile, or a new file of PutRelativeFile) or by desktop Office:
    /// to keep a version, audit, or index. A failure is logged and does not fail the save.
    /// </summary>
    public Func<FileSavedContext, Task>? OnFileSaved { get; set; }
}

/// <summary>
/// A save, given to <see cref="OfficeHostOptions.OnFileSaved"/>.
/// </summary>
/// <param name="File">The file as saved.</param>
/// <param name="Access">The access of the user who saved it.</param>
/// <param name="Editors">The users who edited the file since its last save, as the WOPI client tells them
/// (X-WOPI-Editors), or the user who saved it.</param>
/// <param name="IsAutosave">Whether the WOPI client saved on its own, not because the user asked (Collabora Online's
/// X-COOL-WOPI-IsAutosave); false when unknown.</param>
/// <param name="IsExitSave">Whether the save is that of the user closing the document (X-COOL-WOPI-IsExitSave).</param>
public record FileSavedContext(OfficeFile File, OfficeAccess Access, IReadOnlyList<string> Editors, bool IsAutosave, bool IsExitSave);

/// <summary>
/// The CheckFileInfo response being built, given to <see cref="OfficeHostOptions.OnCheckFileInfo"/>.
/// </summary>
public class CheckFileInfoContext(HttpContext httpContext, OfficeFile file, OfficeAccess access, JsonObject properties)
{
    public HttpContext HttpContext { get; } = httpContext;

    public OfficeFile File { get; } = file;

    public OfficeAccess Access { get; } = access;

    /// <summary>
    /// The properties of the response, as [MS-WOPI] 3.3.5.1.2.2 names them.
    /// </summary>
    public JsonObject Properties { get; } = properties;
}
