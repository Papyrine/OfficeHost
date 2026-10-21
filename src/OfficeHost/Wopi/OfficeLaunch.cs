namespace OfficeHost;

/// <summary>
/// What the page of the application needs to open a file in a WOPI client.
/// </summary>
/// <param name="Url">The URL of the action, to POST the form to.</param>
/// <param name="AccessToken">The access_token form field.</param>
/// <param name="AccessTokenTtl">The access_token_ttl form field: when the token expires, in milliseconds since the Unix epoch.</param>
public record OfficeLaunch(string Url, string AccessToken, long AccessTokenTtl)
{
    /// <summary>
    /// The host page as the WOPI documentation describes it: a form posts the access token to <see cref="Url"/>, in a
    /// full-page iframe. For an application that puts the editor in a page of its own, the form, the iframe and the
    /// script are what it needs.
    /// </summary>
    public string HostPageHtml(string title = "Editor") =>
        $"""
          <!doctype html>
          <html lang="en">
          <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>{WebUtility.HtmlEncode(title)}</title>
          </head>
          <body style="margin: 0">
            <form id="wopi" target="editor" method="post" action="{WebUtility.HtmlEncode(Url)}">
              <input type="hidden" name="access_token" value="{WebUtility.HtmlEncode(AccessToken)}">
              <input type="hidden" name="access_token_ttl" value="{AccessTokenTtl}">
            </form>
            <iframe name="editor" title="{WebUtility.HtmlEncode(title)}" allowfullscreen
                    style="position: fixed; inset: 0; width: 100%; height: 100%; border: none"></iframe>
            <script>document.getElementById('wopi').submit();</script>
          </body>
          </html>
          """;

    /// <summary>
    /// <see cref="HostPageHtml"/> as the result of an endpoint.
    /// </summary>
    public IResult HostPage(string title = "Editor") =>
        Results.Content(HostPageHtml(title), "text/html");
}
