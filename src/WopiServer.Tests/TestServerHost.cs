/// <summary>
/// The whole server running on a <see cref="TestServer"/>, over a temporary local storage.
/// </summary>
sealed class TestServerHost : IAsyncDisposable
{
    public const string IopSecret = "iopsecret";
    public const string WopiSecret = "wopisecret";

    readonly TempFolder folder = new();
    WebApplication app = null!;

    public HttpClient Client { get; private set; } = null!;

    public string StoragePath => folder.Combine("storage");

    public string RecoveryPath => folder.Combine("recovery");

    public StringWriter Log { get; } = new();

    /// <summary>
    /// Bridged apps are reached through this handler; requests to the server itself go to the test server.
    /// </summary>
    public Func<HttpRequestMessage, Task<HttpResponseMessage>>? Apps { get; set; }

    /// <param name="time">The clock of the server, the system one by default.</param>
    /// <param name="multiUser">
    /// Whether each TokenHeader is a distinct user, as with a CS3 storage. The local storage maps all users to root.
    /// </param>
    public static async Task<TestServerHost> Start(string extraConfig = "", TimeProvider? time = null, bool multiUser = false)
    {
        var host = new TestServerHost();
        await host.Init(extraConfig, time ?? TimeProvider.System, multiUser);
        return host;
    }

    async Task Init(string extraConfig, TimeProvider time, bool multiUser)
    {
        Directory.CreateDirectory(StoragePath);
        await File.WriteAllTextAsync(folder.Combine("wopisecret"), $"{WopiSecret}\n");
        await File.WriteAllTextAsync(folder.Combine("iopsecret"), $"{IopSecret}\n");
        var config = Helpers.Ini(
            $"""
             [general]
             storagetype = local
             port = 8880
             wopiurl = http://localhost/
             loglevel = Debug
             homepath = /home/username
             {extraConfig}
             [security]
             usehttps = no
             wopisecretfile = {folder.Combine("wopisecret")}
             iopsecretfile = {folder.Combine("iopsecret")}
             [io]
             chunksize = 4096
             recoverypath = {RecoveryPath}
             [local]
             storagehomepath = {StoragePath}
             """);
        var router = new Router(this);
        IStorage? storage = null;
        if (multiUser)
        {
            storage = new MultiUserStorage(new LocalStorage(new(config), NullLogger.Instance, time));
        }

        app = await WopiApp.Build(
            config,
            new()
            {
                Storage = storage,
                Time = time,
                LogWriter = TextWriter.Synchronized(Log),
                BridgeHandler = router,
                ConfigureBuilder = _ => _.WebHost.UseTestServer()
            });
        await app.StartAsync();
        router.Server = app.GetTestServer().CreateHandler();
        Client = app.GetTestClient();
    }

    public T Service<T>()
        where T : notnull =>
        app.Services.GetRequiredService<T>();

    public string PathOf(string path) =>
        Path.Combine(StoragePath, path.TrimStart('/'));

    public async ValueTask DisposeAsync()
    {
        await app.StopAsync();
        await app.DisposeAsync();
        folder.Dispose();
    }

    /// <summary>
    /// Opens a file through /wopi/iop/openinapp, returning the response.
    /// </summary>
    public async Task<HttpResponseMessage> OpenInApp(string fileId, string viewMode = "VIEW_MODE_READ_WRITE", string appName = "MSOffice", string appUrl = "https://office/edit", string userName = "Alice", string userType = "regular", string userId = "alice")
    {
        var query = $"fileid={Uri.EscapeDataString(fileId)}&viewmode={viewMode}&appname={appName}&appurl={Uri.EscapeDataString(appUrl)}" +
                    $"&appviewurl={Uri.EscapeDataString("https://office/view")}&username={userName}&userid={userId}&usertype={userType}&folderurl={Uri.EscapeDataString("https://efss/files/home")}";
        using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/iop/openinapp?{query}");
        request.Headers.Add("Authorization", $"Bearer {IopSecret}");
        request.Headers.Add("TokenHeader", userId);
        return await Client.SendAsync(request);
    }

    /// <summary>
    /// Opens a file and returns its WOPI file id and access token.
    /// </summary>
    public async Task<(string FileId, string Token, JsonObject Response)> Open(string fileId, string viewMode = "VIEW_MODE_READ_WRITE", string appName = "MSOffice", string userId = "alice", string userType = "regular")
    {
        using var response = await OpenInApp(fileId, viewMode, appName, userType: userType, userId: userId);
        var body = await response.Content.ReadAsStringAsync();
        if (response.StatusCode != HttpStatusCode.OK)
        {
            throw new($"openinapp failed: {response.StatusCode} {body}");
        }

        var json = JsonNode.Parse(body)!.AsObject();
        var appUrl = (string) json["app-url"]!;
        var wopiSrc = Uri.UnescapeDataString(appUrl[(appUrl.IndexOf("WOPISrc=", StringComparison.Ordinal) + 8)..].Split('&')[0]);
        return (wopiSrc.Split('/')[^1], (string) json["form-parameters"]!["access_token"]!, json);
    }

    public Task<HttpResponseMessage> Post(string fileId, string token, string operation, params (string Name, string Value)[] headers) =>
        Send($"/wopi/files/{fileId}", token, null, [("X-WOPI-Override", operation), .. headers]);

    public Task<HttpResponseMessage> PutFile(string fileId, string token, string content, params (string Name, string Value)[] headers) =>
        Send($"/wopi/files/{fileId}/contents", token, content, headers);

    async Task<HttpResponseMessage> Send(string path, string token, string? content, (string Name, string Value)[] headers)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"{path}?access_token={token}");
        request.Content = new StringContent(content ?? "");
        foreach (var (name, value) in headers)
        {
            request.Headers.TryAddWithoutValidation(name, value);
        }

        return await Client.SendAsync(request);
    }

    sealed class Router(TestServerHost host) : HttpMessageHandler
    {
        public HttpMessageHandler? Server { get; set; }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancellationToken)
        {
            if (request.RequestUri!.Host == "localhost")
            {
                using var invoker = new HttpMessageInvoker(Server!, false);
                return await invoker.SendAsync(request, cancellationToken);
            }

            if (host.Apps == null)
            {
                throw new HttpRequestException($"No app at {request.RequestUri}");
            }

            return await host.Apps(request);
        }
    }
}
