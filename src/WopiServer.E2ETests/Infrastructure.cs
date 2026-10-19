using System.IO.Compression;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Xml.Linq;

/// <summary>
/// A WOPI client running in Docker.
/// </summary>
sealed class EditorContainer(string name, int port) : IAsyncDisposable
{
    /// <summary>
    /// The port of the editor on this machine.
    /// </summary>
    public int Port { get; } = port;

    public string Url => $"http://localhost:{Port}";

    /// <summary>
    /// Starts a container and waits until its WOPI discovery answers.
    /// </summary>
    public static async Task<EditorContainer> Start(string image, int port, int containerPort, IReadOnlyDictionary<string, string> environment, TimeSpan startTimeout, Cancel cancel)
    {
        var name = $"wopiserver-e2e-{Guid.NewGuid().ToString("N")[..8]}";
        List<string> args =
        [
            "run", "--detach", "--rm", "--name", name, "--publish", $"{port}:{containerPort}",
            // on Linux, host.docker.internal is not wired by default: map it to the host
            "--add-host", "host.docker.internal:host-gateway"
        ];
        foreach (var (key, value) in environment)
        {
            args.AddRange(["--env", $"{key}={value}"]);
        }

        args.Add(image);
        var (code, output) = await Docker(args, cancel);
        if (code != 0)
        {
            throw new($"docker run failed: {output}");
        }

        var container = new EditorContainer(name, port);
        using var http = new HttpClient();
        var deadline = DateTime.UtcNow + startTimeout;
        while (true)
        {
            try
            {
                using var response = await http.GetAsync($"{container.Url}/hosting/discovery", cancel);
                if (response.IsSuccessStatusCode)
                {
                    return container;
                }
            }
            catch (HttpRequestException)
            {
            }

            if (DateTime.UtcNow > deadline)
            {
                var logs = await container.Logs(cancel);
                await container.DisposeAsync();
                throw new($"{image} did not start in {startTimeout}: {logs}");
            }

            await Task.Delay(2000, cancel);
        }
    }

    public async Task<string> Logs(Cancel cancel) =>
        (await Docker(["logs", "--tail", "200", name], cancel)).Output;

    /// <summary>
    /// The URL of an action of the editor's discovery, without the optional placeholders like `&lt;ui=UI_LLCC&amp;&gt;`.
    /// </summary>
    public async Task<string> ActionUrl(string action, string extension, Cancel cancel)
    {
        using var http = new HttpClient();
        var discovery = XDocument.Parse(await http.GetStringAsync($"{Url}/hosting/discovery", cancel));
        var urlsrc = discovery.Descendants("action")
            .First(_ => (string?) _.Attribute("name") == action && (string?) _.Attribute("ext") == extension)
            .Attribute("urlsrc")!.Value;
        return Regex.Replace(urlsrc, "<[^>]*>", "");
    }

    public async ValueTask DisposeAsync() =>
        await Docker(["rm", "--force", name], Cancel.None);

    static async Task<(int Code, string Output)> Docker(IEnumerable<string> args, Cancel cancel)
    {
        var start = new ProcessStartInfo("docker")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (var arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync(cancel);
        var error = process.StandardError.ReadToEndAsync(cancel);
        await process.WaitForExitAsync(cancel);
        return (process.ExitCode, $"{await output}{await error}");
    }
}

/// <summary>
/// The server on a real port, reachable from the containers as host.docker.internal, over a temporary storage.
/// </summary>
sealed class RunningServer(WebApplication app, StringWriter log, string folder, int port) : IAsyncDisposable
{
    public const string IopSecret = "e2eiopsecret";

    public string Folder { get; } = folder;

    public int Port { get; } = port;

    public string StoragePath => Path.Combine(Folder, "storage");

    public string Log
    {
        get
        {
            lock (log)
            {
                return log.ToString();
            }
        }
    }

    public static async Task<RunningServer> Start(string extraSecurity = "")
    {
        var folder = Directory.CreateTempSubdirectory("wopiserver-e2e").FullName;
        Directory.CreateDirectory(Path.Combine(folder, "storage"));
        await File.WriteAllTextAsync(Path.Combine(folder, "wopisecret"), "e2ewopisecret");
        await File.WriteAllTextAsync(Path.Combine(folder, "iopsecret"), IopSecret);
        var port = FreePort();
        var ini = $"""
                   [general]
                   storagetype = local
                   port = {port}
                   wopiurl = http://host.docker.internal:{port}
                   loglevel = Debug
                   [security]
                   usehttps = no
                   wopisecretfile = {Path.Combine(folder, "wopisecret")}
                   iopsecretfile = {Path.Combine(folder, "iopsecret")}
                   {extraSecurity}
                   [io]
                   recoverypath = {Path.Combine(folder, "recovery")}
                   [local]
                   storagehomepath = {Path.Combine(folder, "storage")}
                   """;
        var configuration = new ConfigurationBuilder()
            .AddIniStream(new MemoryStream(Encoding.UTF8.GetBytes(ini)))
            .Build();
        var log = new StringWriter();
        var app = await WopiApp.Build(
            configuration,
            new()
            {
                LogWriter = TextWriter.Synchronized(log)
            });
        await app.StartAsync();
        return new(app, log, folder, port);
    }

    static int FreePort()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return ((IPEndPoint) listener.LocalEndpoint).Port;
    }

    /// <summary>
    /// Opens a file through /wopi/iop/openinapp: the URL of the editor, and the access token to post to it.
    /// </summary>
    public async Task<(string AppUrl, string Token)> OpenInApp(string file, string appName, string appUrl)
    {
        using var http = new HttpClient();
        using var request = new HttpRequestMessage(HttpMethod.Get, $"http://localhost:{Port}/wopi/iop/openinapp?fileid={Uri.EscapeDataString(file)}&viewmode=VIEW_MODE_READ_WRITE&appname={appName}&appurl={Uri.EscapeDataString(appUrl)}&username=E2E&userid=e2e");
        request.Headers.Add("Authorization", $"Bearer {IopSecret}");
        request.Headers.Add("TokenHeader", "x");
        using var response = await http.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new($"openinapp failed: {response.StatusCode} {body}");
        }

        var json = JsonNode.Parse(body)!;
        return ((string) json["app-url"]!, (string) json["form-parameters"]!["access_token"]!);
    }

    /// <summary>
    /// Waits until the log has a line with the given message.
    /// </summary>
    public async Task WaitForLog(string message, TimeSpan timeout, Cancel cancel)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (!Log.Contains($"\"msg\":\"{message}\""))
        {
            if (DateTime.UtcNow > deadline)
            {
                throw new TimeoutException($"No \"{message}\" in the server log within {timeout}");
            }

            await Task.Delay(500, cancel);
        }
    }

    public async ValueTask DisposeAsync()
    {
        await app.StopAsync();
        await app.DisposeAsync();
        try
        {
            Directory.Delete(Folder, true);
        }
        catch (IOException)
        {
        }
    }
}

/// <summary>
/// A minimal Word document.
/// </summary>
static class Docx
{
    public static byte[] Create(string text)
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, true))
        {
            Add(
                zip,
                "[Content_Types].xml",
                """<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""");
            Add(
                zip,
                "_rels/.rels",
                """<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");
            Add(
                zip,
                "word/document.xml",
                $"""<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>""");
        }

        return buffer.ToArray();
    }

    static void Add(ZipArchive zip, string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(content);
    }

    /// <summary>
    /// The text of a Word document, or null when the file is not one (yet), as while it is being written.
    /// </summary>
    public static string? Text(string path)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            using var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
            var document = XDocument.Parse(reader.ReadToEnd());
            return string.Concat(document.Descendants().Where(_ => _.Name.LocalName == "t").Select(_ => _.Value));
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or System.Xml.XmlException or NullReferenceException)
        {
            return null;
        }
    }
}

/// <summary>
/// Opens an editor the way a WOPI host page does: by posting the access token to the editor URL.
/// </summary>
static class EditorPage
{
    public static async Task Open(IPage page, string appUrl, string token)
    {
        var html = $"""
                    <html><body>
                    <form id="wopi" method="post" action="{WebUtility.HtmlEncode(appUrl)}">
                      <input type="hidden" name="access_token" value="{WebUtility.HtmlEncode(token)}" />
                      <input type="hidden" name="access_token_ttl" value="0" />
                    </form>
                    <script>document.getElementById('wopi').submit();</script>
                    </body></html>
                    """;
        await page.SetContentAsync(html);
    }
}
