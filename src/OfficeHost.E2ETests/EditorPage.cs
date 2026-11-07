/// <summary>
/// Opens an editor the way the page of an application does: by posting the access token to the action URL.
/// </summary>
static class EditorPage
{
    public static Task Open(IPage page, OfficeLaunch launch)
    {
        var html =
            $"""
             <html><body>
             <form id="wopi" method="post" action="{WebUtility.HtmlEncode(launch.Url)}">
               <input type="hidden" name="access_token" value="{WebUtility.HtmlEncode(launch.AccessToken)}" />
               <input type="hidden" name="access_token_ttl" value="{launch.AccessTokenTtl}" />
             </form>
             <script>document.getElementById('wopi').submit();</script>
             </body></html>
             """;
        return page.SetContentAsync(html);
    }
}