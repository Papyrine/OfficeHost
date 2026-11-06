namespace OfficeHost;

/// <summary>
/// The kind of client a request comes from.
/// </summary>
public enum OfficeClient
{
    /// <summary>
    /// A WOPI client, as Collabora Online or Office for the web, calling the endpoints of
    /// <see cref="OfficeHostExtensions.MapOfficeHost"/>.
    /// </summary>
    Wopi,

    /// <summary>
    /// Desktop Office, calling the sites of <see cref="OfficeHostExtensions.MapFssHttp"/>.
    /// </summary>
    Desktop
}

/// <summary>
/// The user a request acts as: the user the application named in the <see cref="OfficeAccess"/> when it issued the
/// access token of the request, and the client the request comes from. Given to the hooks of
/// <see cref="OfficeHostOptions"/> and to the <see cref="IOfficeLockStore"/>, so that the application never has to find
/// the token of the request itself.
/// </summary>
/// <param name="Id">The <see cref="OfficeAccess.UserId"/> of the token.</param>
/// <param name="Name">The <see cref="OfficeAccess.UserName"/> of the token.</param>
/// <param name="Client">The client the request comes from.</param>
public record OfficeUser(string Id, string? Name, OfficeClient Client)
{
    /// <summary>
    /// The user of an access, as a request of the client acts as.
    /// </summary>
    public static OfficeUser From(OfficeAccess access, OfficeClient client) =>
        new(access.UserId, access.UserName, client);
}
