namespace OfficeHost;

/// <summary>
/// What an access token grants: a user's access to one file.
/// </summary>
/// <param name="FileId">The id of the file, as <see cref="IOfficeStorage"/> knows it.</param>
/// <param name="UserId">A unique and stable id of the user.</param>
public record OfficeAccess(string FileId, string UserId)
{
    /// <summary>
    /// The name of the user, as the WOPI client shows it.
    /// </summary>
    public string? UserName { get; init; }

    /// <summary>
    /// A user that is not signed in, such as the visitor of a public link.
    /// </summary>
    public bool IsAnonymous { get; init; }

    /// <summary>
    /// The user can change the file: lock it, save it, and delete it.
    /// </summary>
    public bool CanWrite { get; init; }

    /// <summary>
    /// The user can create files next to this one, as the Save As of the WOPI clients does (PutRelativeFile), and of
    /// desktop Office, which then edits the new file through the site of the token. Both can also replace an unlocked
    /// file of the chosen name, but not this one, nor read a file before replacing it.
    /// </summary>
    public bool CanWriteRelative { get; init; }

    /// <summary>
    /// The user can rename the file.
    /// </summary>
    public bool CanRename { get; init; }

    /// <summary>
    /// The user can only view the file in the browser: no download, print or copy (RestrictedWebViewOnly). The WOPI
    /// client enforces it, so it holds only for requests that come from a WOPI client: set
    /// <see cref="OfficeHostOptions.RequireProof"/>.
    /// </summary>
    public bool ViewOnly { get; init; }
}

/// <summary>
/// An access token and what it grants.
/// </summary>
/// <param name="Value">The token, as the WOPI client passes it in the access_token parameter.</param>
public record OfficeToken(string Value, OfficeAccess Access, DateTimeOffset Expires);

/// <summary>
/// Issues and validates the access tokens. The built-in implementation signs the access in the token with
/// <see cref="OfficeHostOptions.SigningKey"/>; register another one to use opaque tokens kept in a database, for instance.
/// </summary>
public interface IOfficeTokens
{
    Task<OfficeToken> Issue(OfficeAccess access, Cancel cancel = default);

    /// <summary>
    /// The token, or null when it is invalid or expired.
    /// </summary>
    Task<OfficeToken?> Validate(string token, Cancel cancel = default);
}
