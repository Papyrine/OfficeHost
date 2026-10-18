using System.Security.Cryptography.X509Certificates;

/// <summary>
/// Options to embed or test the server.
/// </summary>
public class WopiAppOptions
{
    /// <summary>
    /// A storage layer to use instead of the one configured by `storagetype`.
    /// </summary>
    public IStorage? Storage { get; init; }

    public TimeProvider? Time { get; init; }

    /// <summary>
    /// Where the log goes, instead of the configured `loghandler`/`logdest`.
    /// </summary>
    public TextWriter? LogWriter { get; init; }

    /// <summary>
    /// The handler of the outgoing requests: of the bridge, to the bridged apps and back to this server, and of the
    /// proof keys validation, to the discovery of the WOPI clients.
    /// </summary>
    public HttpMessageHandler? HttpHandler { get; init; }

    public Action<WebApplicationBuilder>? ConfigureBuilder { get; init; }
}

/// <summary>
/// The Web-application Open Platform Interface (WOPI) gateway for the ScienceMesh IOP.
/// </summary>
public static class WopiApp
{
    /// <summary>
    /// Initialises the application, failing in case of any configuration error, and maps the endpoints.
    /// </summary>
    public static async Task<WebApplication> Build(IConfiguration configuration, WopiAppOptions? options = null)
    {
        options ??= new();
        var config = new WopiConfig(configuration);
        var time = options.Time ?? TimeProvider.System;

        // detect hostname, or take it from the environment if set e.g. by docker
        var hostname = Environment.GetEnvironmentVariable("HOST_HOSTNAME");
        if (string.IsNullOrEmpty(hostname))
        {
            hostname = Environment.MachineName;
        }

        // configure the logging
        var logProvider = new JsonLogProvider(options.LogWriter ?? LogWriter(config), options.LogWriter == null, hostname, () => config.LogLevel);
        var builder = WebApplication.CreateBuilder(
            new WebApplicationOptions
            {
                ContentRootPath = AppContext.BaseDirectory
            });
        builder.Logging.ClearProviders();
        builder.Logging.SetMinimumLevel(LogLevel.Trace);
        builder.Logging.AddProvider(logProvider);
        var loggers = LoggerFactory.Create(
            _ =>
            {
                _.SetMinimumLevel(LogLevel.Trace);
                _.AddProvider(logProvider);
            });
        var log = loggers.CreateLogger("wopiserver");

        WopiState state;
        X509Certificate2? certificate = null;
        Bridge bridge;
        try
        {
            // load the requested storage layer
            var storageType = config.GetRequired("general", "storagetype");
            var storage = options.Storage ?? CreateStorage(storageType, config, loggers.CreateLogger($"{storageType}iface"), time);
            var useHttps = string.Equals(config.GetRequired("security", "usehttps"), "yes", StringComparison.OrdinalIgnoreCase);
            // validate the certificates exist if running in https mode
            if (useHttps)
            {
                try
                {
                    certificate = LoadCertificate(config.GetRequired("security", "wopicert"), config.GetRequired("security", "wopikey"));
                }
                catch (Exception exception) when (exception is IOException or CryptographicException or UnauthorizedAccessException)
                {
                    log.LogError("Failed to open the provided certificate or key to start in https mode");
                    throw;
                }
            }

            var recoveryPath = config.Get("io", "recoverypath", "/var/spool/wopirecovery");
            Directory.CreateDirectory(recoveryPath);

            // WOPI proxy configuration (optional)
            var wopiProxy = config.Get("general", "wopiproxy", "").Trim('/');
            string? wopiProxyKey = null;
            var proxyKeyFile = config.Get("general", "wopiproxysecretfile", "");
            if (proxyKeyFile.Length > 0)
            {
                wopiProxyKey = await ReadSecret(proxyKeyFile);
            }

            var proxiedAppName = config.Get("general", "proxiedappname", "");
            if (proxiedAppName.Length > 0 &&
                (wopiProxy.Length == 0 || wopiProxyKey == null))
            {
                throw new ConfigException("Proxed app configured but missing wopiproxy or shared key");
            }

            state = new()
            {
                Config = config,
                Storage = storage,
                Time = time,
                StorageType = storageType,
                Port = config.GetInt("general", "port"),
                UseHttps = useHttps,
                WopiSecret = await ReadSecret(config.GetRequired("security", "wopisecretfile")),
                IopSecret = await ReadSecret(config.GetRequired("security", "iopsecretfile")),
                WopiUrl = config.GetRequired("general", "wopiurl").Trim('/'),
                HomePath = config.Get("general", "homepath", "/home/username"),
                RecoveryPath = recoveryPath,
                WopiProxy = wopiProxy,
                WopiProxyKey = wopiProxyKey,
                ProxiedAppName = proxiedAppName,
                NonOfficeTypes = Split(config.Get("apps", "nonofficetypes", "")),
                CodeTypes = Split(config.Get("apps", "codeofficetypes", ".odt .ods .odp")),
                EnableRename = config.IsTrueOrYes("general", "enablerename", false),
                ProofKeys = new(config, ProofKeysClient(options.HttpHandler), time, loggers.CreateLogger("proofkeys"))
            };

            // initialize the bridge
            bridge = new(state, new(state, loggers.CreateLogger("wopiutils")), BridgeHttp(config, options.HttpHandler), loggers);
            await bridge.Init();
        }
        catch (Exception exception) when (exception is ConfigException or IOException or UnauthorizedAccessException or CryptographicException or UnsupportedAppException)
        {
            // any error we get here with the configuration is fatal
            log.LogCritical("Failed to initialize the service, aborting {error}", exception.Message);
            logProvider.Dispose();
            throw;
        }

        var utils = new WopiUtils(state, loggers.CreateLogger("wopiutils"));
        var core = new WopiCore(state, utils, loggers.CreateLogger("wopi"));
        var endpoints = new WopiEndpoints(state, utils, core, bridge, log);

        builder.Services.AddSingleton(state);
        builder.Services.AddSingleton(bridge);
        builder.Services.AddHostedService(_ => new BridgeSaveService(bridge, loggers));
        builder.WebHost.ConfigureKestrel(
            kestrel =>
            {
                // Office documents can be large: do not cap the body of PutFile requests
                kestrel.Limits.MaxRequestBodySize = null;
                kestrel.AddServerHeader = false;
                kestrel.RequestHeaderEncodingSelector = _ => Encoding.UTF8;
                kestrel.ResponseHeaderEncodingSelector = _ => Encoding.UTF8;
                kestrel.ListenAnyIP(
                    state.Port,
                    listen =>
                    {
                        if (certificate != null)
                        {
                            listen.UseHttps(certificate);
                        }
                    });
            });
        options.ConfigureBuilder?.Invoke(builder);

        var app = builder.Build();
        app.Lifetime.ApplicationStopped.Register(logProvider.Dispose);
        endpoints.Map(app);
        if (state.UseHttps)
        {
            log.LogInformation("WOPI Server starting in standalone secure mode {port} {wopiurl} {version}", state.Port, state.WopiUrl, WopiState.Version);
        }
        else
        {
            log.LogInformation("WOPI Server starting in unsecure/embedded mode {port} {wopiurl} {version}", state.Port, state.WopiUrl, WopiState.Version);
        }

        if (state.ProofKeys!.Enabled)
        {
            log.LogInformation(
                "Validating X-WOPI-Proof against the proof keys of the WOPI clients {discoveryurls} {required}",
                string.Join(' ', state.ProofKeys.DiscoveryUrls),
                state.ProofKeys.Required);
        }

        return app;
    }

    static TextWriter LogWriter(WopiConfig config)
    {
        if (string.Equals(config.Get("general", "loghandler", "file"), "stream", StringComparison.OrdinalIgnoreCase))
        {
            if (string.Equals(config.Get("general", "logdest", "stdout"), "stdout", StringComparison.OrdinalIgnoreCase))
            {
                return Console.Out;
            }

            return Console.Error;
        }

        var path = config.Get("general", "logdest", "/var/log/wopi/wopiserver.log");
        return new StreamWriter(new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete), new UTF8Encoding(false));
    }

    static IStorage CreateStorage(string storageType, WopiConfig config, ILogger log, TimeProvider time) =>
        storageType switch
        {
            "local" => new LocalStorage(config, log, time),
            "xroot" or "cs3" => throw new ConfigException($"Storage type {storageType} is not supported by this port of the WOPI server"),
            _ => throw new ConfigException($"Unsupported/Unknown storage type {storageType}")
        };

    static X509Certificate2 LoadCertificate(string certificatePath, string keyPath)
    {
        using var pem = X509Certificate2.CreateFromPemFile(certificatePath, keyPath);
        // an ephemeral PEM key is not usable by SslStream on Windows: round trip it through PKCS#12
        return X509CertificateLoader.LoadPkcs12(pem.Export(X509ContentType.Pkcs12), null);
    }

    /// <summary>
    /// The content of a secret file, without the trailing line break, including a Windows one.
    /// </summary>
    static async Task<string> ReadSecret(string path) =>
        (await File.ReadAllTextAsync(path)).Trim('\r', '\n');

    static string[] Split(string value) =>
        value.Split((char[]?) null, StringSplitOptions.RemoveEmptyEntries);

    static BridgeHttp BridgeHttp(WopiConfig config, HttpMessageHandler? handler)
    {
        if (handler != null)
        {
            return new(new(handler, false), new(handler, false));
        }

        var sslVerify = config.IsTrueOrYes("bridge", "sslverify", true);
        return new(Client(sslVerify, true), Client(sslVerify, false));
    }

    /// <summary>
    /// The discovery of the WOPI clients is fetched with full certificate validation: the keys are what is trusted.
    /// </summary>
    static HttpClient ProofKeysClient(HttpMessageHandler? handler)
    {
        if (handler != null)
        {
            return new(handler, false);
        }

        return Client(true, true);
    }

    static HttpClient Client(bool sslVerify, bool allowRedirect)
    {
        var handler = new SocketsHttpHandler
        {
            AllowAutoRedirect = allowRedirect
        };
        if (!sslVerify)
        {
            handler.SslOptions.RemoteCertificateValidationCallback = (_, _, _, _) => true;
        }

        return new(handler)
        {
            Timeout = TimeSpan.FromSeconds(10)
        };
    }
}
