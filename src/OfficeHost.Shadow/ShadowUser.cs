namespace OfficeHost;

/// <summary>
/// The user of OfficeHost the shadow replays a user of SharePoint as.
/// </summary>
public record ShadowUser(string Id, string? Name)
{
    /// <summary>
    /// The user of a login of SharePoint: the whole login as id (<c>i:0#.w|contoso\alice</c>), the account without its
    /// domain or claim as name (<c>alice</c>). A request without a user (SharePoint serves IsOnlyClient without sign-in)
    /// is <c>anonymous</c>.
    /// </summary>
    public static ShadowUser FromLogin(string login)
    {
        if (login.Length == 0)
        {
            return new("anonymous", null);
        }

        return new(login, login.Split('|')[^1].Split('\\')[^1]);
    }
}
