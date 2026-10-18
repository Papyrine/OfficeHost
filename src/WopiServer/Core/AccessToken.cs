/// <summary>
/// File view mode: reference is `ViewMode` at
/// https://github.com/cs3org/cs3apis/blob/master/cs3/app/provider/v1beta1/provider_api.proto
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<ViewMode>))]
public enum ViewMode
{
    /// <summary>
    /// The file can be opened but not downloaded.
    /// </summary>
    [JsonStringEnumMemberName("VIEW_MODE_VIEW_ONLY")]
    ViewOnly,

    /// <summary>
    /// The file can be downloaded.
    /// </summary>
    [JsonStringEnumMemberName("VIEW_MODE_READ_ONLY")]
    ReadOnly,

    /// <summary>
    /// The file can be downloaded and updated, and the app should be shown in edit mode.
    /// </summary>
    [JsonStringEnumMemberName("VIEW_MODE_READ_WRITE")]
    ReadWrite,

    /// <summary>
    /// The file can be downloaded and updated, and the app should be shown in preview mode.
    /// </summary>
    [JsonStringEnumMemberName("VIEW_MODE_PREVIEW")]
    Preview,

    /// <summary>
    /// The file can be viewed in embedded mode, no editing is allowed.
    /// </summary>
    [JsonStringEnumMemberName("VIEW_MODE_EMBEDDED")]
    Embedded
}

/// <summary>
/// App user types as given by https://github.com/cs3org/reva/blob/master/pkg/app/provider/wopi/wopi.go
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter<UserType>))]
public enum UserType
{
    [JsonStringEnumMemberName("invalid")]
    Invalid,

    /// <summary>
    /// Regular user, logged in the local ID provider.
    /// </summary>
    [JsonStringEnumMemberName("regular")]
    Regular,

    /// <summary>
    /// Federated/external user, logged in the local ID provider but with no home space.
    /// </summary>
    [JsonStringEnumMemberName("federated")]
    Federated,

    /// <summary>
    /// OCM user, logged in a remote ID provider.
    /// </summary>
    [JsonStringEnumMemberName("ocm")]
    Ocm,

    /// <summary>
    /// Anonymous user, accessing a public link.
    /// </summary>
    [JsonStringEnumMemberName("anonymous")]
    Anonymous
}

static class WopiEnums
{
    public static string Value(this ViewMode mode) =>
        mode switch
        {
            ViewMode.ViewOnly => "VIEW_MODE_VIEW_ONLY",
            ViewMode.ReadOnly => "VIEW_MODE_READ_ONLY",
            ViewMode.ReadWrite => "VIEW_MODE_READ_WRITE",
            ViewMode.Preview => "VIEW_MODE_PREVIEW",
            _ => "VIEW_MODE_EMBEDDED"
        };

    public static ViewMode? ParseViewMode(string? value) =>
        value switch
        {
            "VIEW_MODE_VIEW_ONLY" => ViewMode.ViewOnly,
            "VIEW_MODE_READ_ONLY" => ViewMode.ReadOnly,
            "VIEW_MODE_READ_WRITE" => ViewMode.ReadWrite,
            "VIEW_MODE_PREVIEW" => ViewMode.Preview,
            "VIEW_MODE_EMBEDDED" => ViewMode.Embedded,
            _ => null
        };

    public static string Value(this UserType type) =>
        type switch
        {
            UserType.Regular => "regular",
            UserType.Federated => "federated",
            UserType.Ocm => "ocm",
            UserType.Anonymous => "anonymous",
            _ => "invalid"
        };

    public static UserType? ParseUserType(string? value) =>
        value switch
        {
            "invalid" => UserType.Invalid,
            "regular" => UserType.Regular,
            "federated" => UserType.Federated,
            "ocm" => UserType.Ocm,
            "anonymous" => UserType.Anonymous,
            _ => null
        };

    /// <summary>
    /// The modes that allow writing to the file.
    /// </summary>
    public static bool CanWrite(this ViewMode mode) =>
        mode is ViewMode.ReadWrite or ViewMode.Preview;

    /// <summary>
    /// The modes that do not allow downloading the file.
    /// </summary>
    public static bool IsViewOnly(this ViewMode mode) =>
        mode is ViewMode.ViewOnly or ViewMode.Embedded;
}

/// <summary>
/// The payload of a WOPI access token. Property names and order match the Python implementation, so tokens
/// are interchangeable between the two given the same secret.
/// </summary>
public class AccessToken
{
    [JsonPropertyName("userid")]
    public required string UserId { get; init; }

    /// <summary>
    /// In the form `username!userid_in_stat_format`.
    /// </summary>
    [JsonPropertyName("wopiuser")]
    public required string WopiUser { get; init; }

    [JsonPropertyName("usertype")]
    public required UserType UserType { get; init; }

    [JsonPropertyName("filename")]
    public required string FileName { get; init; }

    [JsonPropertyName("fileid")]
    public required string FileId { get; init; }

    [JsonPropertyName("username")]
    public required string UserName { get; init; }

    [JsonPropertyName("viewmode")]
    public required ViewMode ViewMode { get; init; }

    [JsonPropertyName("folderurl")]
    public required string FolderUrl { get; init; }

    [JsonPropertyName("endpoint")]
    public required string Endpoint { get; init; }

    [JsonPropertyName("appname")]
    public required string AppName { get; init; }

    [JsonPropertyName("trace")]
    public string? Trace { get; init; }

    [JsonPropertyName("exp")]
    public required long Exp { get; init; }

    [JsonPropertyName("iss")]
    public required string Iss { get; init; }

    [JsonPropertyName("appediturl")]
    public string? AppEditUrl { get; init; }

    [JsonPropertyName("appviewurl")]
    public string? AppViewUrl { get; init; }

    /// <summary>
    /// The username part of <see cref="WopiUser"/>.
    /// </summary>
    [JsonIgnore]
    public string WopiUserName => WopiUser.Split('!')[0];

    /// <summary>
    /// The userid part of <see cref="WopiUser"/>.
    /// </summary>
    [JsonIgnore]
    public string WopiUserId => WopiUser.Split('!')[^1];
}
