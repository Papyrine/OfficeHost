namespace OfficeHost;

/// <summary>
/// What the page of the application needs to open a file in a WOPI client.
/// </summary>
/// <param name="Url">The URL of the action, to POST the form to.</param>
/// <param name="AccessToken">The access_token form field.</param>
/// <param name="AccessTokenTtl">The access_token_ttl form field: when the token expires, in milliseconds since the Unix epoch.</param>
public record OfficeLaunch(string Url, string AccessToken, long AccessTokenTtl);