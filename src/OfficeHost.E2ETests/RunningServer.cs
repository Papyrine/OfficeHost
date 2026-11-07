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
    public static async Task<RunningServer> Start(Action<OfficeHostOptions>? configure = null, string host = "host.docker.internal")
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
        builder.Services.AddSingleton<IOfficeStorage>(storage);
        builder.Services.AddSingleton<IOfficeUserInfoStore, MemoryUserInfoStore>();
        builder.Services.AddOfficeHost(
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
        app.MapOfficeHost();
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
        var token = await Service<IOfficeTokens>().Issue(Access(fileId));
        return (fileId, token.Value);
    }

    public static OfficeAccess Access(string fileId) =>
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
    public async Task<(string FileId, OfficeLaunch Launch)> Launch(string path, string discoveryUrl)
    {
        var fileId = Storage.GetFileId(path);
        return (fileId, await Service<OfficeLauncher>().Launch(Access(fileId), "edit", discoveryUrl: discoveryUrl));
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