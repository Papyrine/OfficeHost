/// <summary>
/// The pages of the sample.
/// </summary>
static class SamplePages
{
    static string Page(string title, string body) =>
        $$"""
          <!doctype html>
          <html lang="en">
          <head>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>{{WebUtility.HtmlEncode(title)}}</title>
            <style>
              body { font-family: system-ui, sans-serif; margin: 0; }
              main { max-width: 48rem; margin: 2rem auto; padding: 0 1rem; }
              td { padding: 0.25rem 1rem 0.25rem 0; }
              code { overflow-wrap: anywhere; }
              pre { white-space: pre-wrap; overflow-wrap: anywhere; }
            </style>
          </head>
          <body>{{body}}</body>
          </html>
          """;

    /// <summary>
    /// The files of the storage, with links to open them.
    /// </summary>
    public static string Files(FileSystemStorage storage)
    {
        var rows = new StringBuilder();
        foreach (var path in Directory.EnumerateFiles(storage.Root).Order(StringComparer.OrdinalIgnoreCase))
        {
            var name = Path.GetFileName(path);
            var id = Uri.EscapeDataString(storage.GetFileId(name));
            rows.Append(
                $"""
                 <tr>
                   <td>{WebUtility.HtmlEncode(name)}</td>
                   <td><a href="/edit/{id}">Edit</a></td>
                   <td><a href="/view/{id}">View</a></td>
                   <td><a href="/desktop/{id}">Open in desktop Office</a></td>
                 </tr>
                 """);
        }

        return Page(
            "Files",
            $"""
             <main>
               <h1>Files</h1>
               <p>The files of <code>{WebUtility.HtmlEncode(storage.Root)}</code>. Add files to the folder to see them here.</p>
               <table>{rows}</table>
             </main>
             """);
    }

    /// <summary>
    /// The page that holds the WOPI client: a form posts the access token to the action URL, in a full-page iframe, as
    /// the WOPI documentation of the host page describes.
    /// </summary>
    public static string Editor(OfficeLaunch launch) =>
        Page(
            "Editor",
            $"""
             <form id="wopi" target="editor" method="post" action="{WebUtility.HtmlEncode(launch.Url)}">
               <input type="hidden" name="access_token" value="{WebUtility.HtmlEncode(launch.AccessToken)}">
               <input type="hidden" name="access_token_ttl" value="{launch.AccessTokenTtl}">
             </form>
             <iframe name="editor" title="Editor" allowfullscreen
                     style="position: fixed; inset: 0; width: 100%; height: 100%; border: none"></iframe>
             <script>document.getElementById('wopi').submit();</script>
             """);

    /// <summary>
    /// The error page when the WOPI client cannot be reached, with how to start Collabora Online.
    /// </summary>
    public static string Unreachable(string message) =>
        Page(
            "Cannot open the file",
            $"""
             <main>
               <h1>Cannot open the file</h1>
               <p>{WebUtility.HtmlEncode(message)}</p>
               <p>Is Collabora Online running? Start it in Docker, as the Sample section of the readme shows:</p>
               <pre>{WebUtility.HtmlEncode(collaboraCommand)}</pre>
               <p>Then try again: the sample fetches the discovery of Collabora as soon as it answers.</p>
               <p><a href="/">Back to the files</a></p>
             </main>
             """);

    const string collaboraCommand =
        """docker run --rm --publish 9980:9980 --add-host host.docker.internal:host-gateway --env "domain=host\.docker\.internal.*" --env "extra_params=--o:ssl.enable=false --o:ssl.termination=false --o:home_mode.enable=true" collabora/code""";

    /// <summary>
    /// The link that opens a file in desktop Office, followed at once, and offered in case the browser asks first.
    /// </summary>
    public static string Desktop(string link) =>
        Page(
            "Open in desktop Office",
            $"""
             <main>
               <h1>Open in desktop Office</h1>
               <p><a id="open" href="{WebUtility.HtmlEncode(link)}">Open the file in desktop Office</a></p>
               <p><a href="/">Back to the files</a></p>
             </main>
             <script>document.getElementById("open").click()</script>
             """);

    public static string Message(string message) =>
        Page(
            "Cannot open the file",
            $"""
             <main>
               <h1>Cannot open the file</h1>
               <p>{WebUtility.HtmlEncode(message)}</p>
               <p><a href="/">Back to the files</a></p>
             </main>
             """);
}