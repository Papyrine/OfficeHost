using System.IO.Compression;
using System.Net.Sockets;

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

    public string DiscoveryUrl => $"{Url}/hosting/discovery";

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
/// An application hosting the WOPI endpoints on a real port, over a temporary <see cref="FileSystemStorage"/>. The
/// containers reach it as host.docker.internal.
/// </summary>
sealed class RunningServer(WebApplication app, string folder, FileSystemStorage storage, List<string> requests, TestLog log) : IAsyncDisposable
{
    public string Folder { get; } = folder;

    public FileSystemStorage Storage { get; } = storage;

    public string StoragePath => Storage.Root;

    public TestLog Log { get; } = log;

    /// <summary>
    /// The WOPI requests served, as `GET /wopi/files/{id}/contents 200`.
    /// </summary>
    public IReadOnlyList<string> Requests
    {
        get
        {
            lock (requests)
            {
                return requests.ToList();
            }
        }
    }

    /// <param name="host">The host name the WOPI clients reach this machine by.</param>
    public static async Task<RunningServer> Start(Action<WopiOptions>? configure = null, string host = "host.docker.internal")
    {
        var folder = Directory.CreateTempSubdirectory("wopiserver-e2e").FullName;
        var storage = new FileSystemStorage(Directory.CreateDirectory(Path.Combine(folder, "storage")).FullName);
        var port = FreePort();
        var builder = WebApplication.CreateSlimBuilder();
        // the containers reach the host through its gateway address, not the loopback one
        builder.WebHost.UseUrls($"http://0.0.0.0:{port}");
        var log = new TestLog();
        builder.Logging.ClearProviders();
        builder.Logging.SetMinimumLevel(LogLevel.Debug);
        builder.Logging.AddProvider(log);
        builder.Services.AddSingleton<IWopiStorage>(storage);
        builder.Services.AddSingleton<IWopiUserInfoStore, MemoryUserInfoStore>();
        builder.Services.AddWopi(
            _ =>
            {
                _.SigningKey = "the signing key of the end-to-end tests";
                _.BaseUrl = $"http://{host}:{port}";
                // the pages of the application that open a file, which PutRelativeFile returns for the new file
                _.HostEditUrl = (file, _) => $"http://{host}:{port}/edit/{file.Id}";
                _.HostViewUrl = (file, _) => $"http://{host}:{port}/view/{file.Id}";
                configure?.Invoke(_);
            });
        var app = builder.Build();
        var requests = new List<string>();
        app.Use(
            async (context, next) =>
            {
                await next(context);
                lock (requests)
                {
                    requests.Add($"{context.Request.Method} {context.Request.Path} {context.Response.StatusCode}");
                }
            });
        app.MapWopi();
        await app.StartAsync();
        return new(app, folder, storage, requests, log);
    }

    static int FreePort()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return ((IPEndPoint) listener.LocalEndpoint).Port;
    }

    public T Service<T>()
        where T : notnull =>
        app.Services.GetRequiredService<T>();

    /// <summary>
    /// The id of a file of the storage, and an access token to it that allows everything.
    /// </summary>
    public async Task<(string FileId, string Token)> Open(string path)
    {
        var fileId = Storage.GetFileId(path);
        var token = await Service<IWopiTokens>().Issue(Access(fileId));
        return (fileId, token.Value);
    }

    public static WopiAccess Access(string fileId) =>
        new(fileId, "e2e")
        {
            UserName = "E2E",
            CanWrite = true,
            CanWriteRelative = true,
            CanRename = true
        };

    /// <summary>
    /// Launches the edit action of a file in the WOPI client of the discovery.
    /// </summary>
    public async Task<(string FileId, WopiLaunch Launch)> Launch(string path, string discoveryUrl)
    {
        var fileId = Storage.GetFileId(path);
        return (fileId, await Service<WopiLauncher>().Launch(Access(fileId), "edit", discoveryUrl: discoveryUrl));
    }

    /// <summary>
    /// Waits until a request matching the given one was served, as `POST /wopi/files/{id}/contents 200`.
    /// </summary>
    public async Task WaitForRequest(string request, TimeSpan timeout, Cancel cancel)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (!Requests.Contains(request))
        {
            if (DateTime.UtcNow > deadline)
            {
                throw new TimeoutException($"No {request} within {timeout}, served: {string.Join(", ", Requests)}");
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
/// The user info in memory.
/// </summary>
sealed class MemoryUserInfoStore : IWopiUserInfoStore
{
    readonly ConcurrentDictionary<string, string> values = new();

    public Task<string?> Get(string userId, Cancel cancel = default) =>
        Task.FromResult(values.GetValueOrDefault(userId));

    public Task Set(string userId, string userInfo, Cancel cancel = default)
    {
        values[userId] = userInfo;
        return Task.CompletedTask;
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
/// Opens an editor the way the page of an application does: by posting the access token to the action URL.
/// </summary>
static class EditorPage
{
    public static Task Open(IPage page, WopiLaunch launch)
    {
        var html = $"""
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
