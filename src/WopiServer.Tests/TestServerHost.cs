/// <summary>
/// The access a test opens a file with.
/// </summary>
public enum Mode
{
    /// <summary>
    /// Write, write relative and rename.
    /// </summary>
    Edit,

    ReadOnly,

    /// <summary>
    /// Read only, and no download.
    /// </summary>
    ViewOnly
}

/// <summary>
/// An application with the WOPI endpoints on a <see cref="TestServer"/>, over a temporary <see cref="FileSystemStorage"/>.
/// </summary>
sealed class TestServerHost : IAsyncDisposable
{
    public const string SigningKey = "the signing key of the tests, long enough";

    readonly TempFolder folder = new();
    WebApplication app = null!;

    public HttpClient Client { get; private set; } = null!;

    public string StoragePath => folder.Combine("storage");

    public FileSystemStorage Storage { get; private set; } = null!;

    public TestLog Log { get; } = new();

    /// <summary>
    /// Answers the outgoing requests, as those of the WOPI discovery.
    /// </summary>
    public Func<HttpRequestMessage, Task<HttpResponseMessage>>? Http { get; set; }

    /// <param name="configure">Changes the options, after the defaults of the tests.</param>
    /// <param name="time">The clock of the application, the system one by default.</param>
    /// <param name="wrapStorage">Decorates the storage, to simulate storage behaviours.</param>
    /// <param name="userInfo">Whether an <see cref="IWopiUserInfoStore"/> is registered.</param>
    /// <param name="kestrel">Runs on Kestrel, on a free local port, rather than on a TestServer.</param>
    public static async Task<TestServerHost> Start(Action<WopiOptions>? configure = null, TimeProvider? time = null, Func<IWopiStorage, IWopiStorage>? wrapStorage = null, bool userInfo = true, Action<IServiceCollection>? services = null, bool kestrel = false)
    {
        var host = new TestServerHost();
        await host.Init(configure, time, wrapStorage, userInfo, services, kestrel);
        return host;
    }

    async Task Init(Action<WopiOptions>? configure, TimeProvider? time, Func<IWopiStorage, IWopiStorage>? wrapStorage, bool userInfo, Action<IServiceCollection>? configureServices, bool kestrel)
    {
        Directory.CreateDirectory(StoragePath);
        Storage = new(StoragePath);
        var builder = WebApplication.CreateSlimBuilder();
        if (kestrel)
        {
            builder.WebHost.UseUrls("http://127.0.0.1:0");
        }
        else
        {
            builder.WebHost.UseTestServer();
        }

        builder.Logging.ClearProviders();
        builder.Logging.SetMinimumLevel(LogLevel.Debug);
        builder.Logging.AddProvider(Log);
        var services = builder.Services;
        if (time != null)
        {
            services.AddSingleton(time);
        }

        services.AddSingleton<IWopiStorage>(wrapStorage?.Invoke(Storage) ?? Storage);
        if (userInfo)
        {
            services.AddSingleton<IWopiUserInfoStore, MemoryUserInfoStore>();
        }

        services.AddWopi(
            options =>
            {
                options.SigningKey = SigningKey;
                options.BaseUrl = "http://localhost";
                options.HostEditUrl = (file, _) => $"https://app/edit/{file.Id}";
                options.HostViewUrl = (file, _) => $"https://app/view/{file.Id}";
                configure?.Invoke(options);
            });
        services.AddHttpClient(WopiOptions.HttpClientName)
            .ConfigurePrimaryHttpMessageHandler(() => new Handler(this));
        configureServices?.Invoke(services);
        app = builder.Build();
        app.MapWopi();
        await app.StartAsync();
        if (kestrel)
        {
            var address = app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()!.Addresses.First();
            Client = new()
            {
                BaseAddress = new(address)
            };
            return;
        }

        Client = app.GetTestClient();
    }

    public T Service<T>()
        where T : notnull =>
        app.Services.GetRequiredService<T>();

    public string PathOf(string path) =>
        Path.Combine(StoragePath, path.TrimStart('/'));

    public async ValueTask DisposeAsync()
    {
        Client.Dispose();
        await app.StopAsync();
        await app.DisposeAsync();
        folder.Dispose();
    }

    public static WopiAccess Access(string fileId, Mode mode = Mode.Edit, string userId = "alice", bool anonymous = false) =>
        new(fileId, userId)
        {
            UserName = anonymous ? null : char.ToUpperInvariant(userId[0]) + userId[1..],
            IsAnonymous = anonymous,
            CanWrite = mode == Mode.Edit,
            CanWriteRelative = mode == Mode.Edit,
            CanRename = mode == Mode.Edit,
            ViewOnly = mode == Mode.ViewOnly
        };

    /// <summary>
    /// The id of a file of the storage, and an access token to it.
    /// </summary>
    public async Task<(string FileId, string Token)> Open(string path, Mode mode = Mode.Edit, string userId = "alice", bool anonymous = false)
    {
        var fileId = Storage.GetFileId(path.TrimStart('/'));
        return (fileId, await Token(Access(fileId, mode, userId, anonymous)));
    }

    public async Task<string> Token(WopiAccess access) =>
        (await Service<IWopiTokens>().Issue(access)).Value;

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

    sealed class Handler(TestServerHost host) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancellationToken)
        {
            if (host.Http == null)
            {
                throw new HttpRequestException($"Unreachable {request.RequestUri}");
            }

            return host.Http(request);
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
/// The log lines of an application, as `Level Category: message`.
/// </summary>
sealed class TestLog : ILoggerProvider
{
    readonly List<string> lines = [];

    public IReadOnlyList<string> Lines
    {
        get
        {
            lock (lines)
            {
                return lines.ToList();
            }
        }
    }

    public override string ToString() =>
        string.Join('\n', Lines);

    public ILogger CreateLogger(string categoryName) =>
        new Logger(this, categoryName);

    public void Dispose()
    {
    }

    sealed class Logger(TestLog log, string category) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull =>
            null;

        public bool IsEnabled(LogLevel logLevel) =>
            true;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception, Func<TState, Exception?, string> formatter)
        {
            lock (log.lines)
            {
                log.lines.Add($"{logLevel} {category}: {formatter(state, exception)}");
            }
        }
    }
}
