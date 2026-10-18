using System.Net.Sockets;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

/// <summary>
/// The startup of the server: configuration errors, HTTPS, the log destinations, the WOPI proxy and the command line.
/// </summary>
public class StartupTests
{
    /// <summary>
    /// A folder with the secret files and a storage, and the configuration over them.
    /// </summary>
    sealed class Setup : IDisposable
    {
        public TempFolder Folder { get; } = new();

        public Setup()
        {
            Directory.CreateDirectory(Folder.Combine("storage"));
            File.WriteAllText(Folder.Combine("wopisecret"), "wopisecret\r\n");
            File.WriteAllText(Folder.Combine("iopsecret"), "iopsecret\n");
        }

        /// <param name="general">Lines of the [general] section, replacing the storage type and URL when given.</param>
        public string Config(string general = "storagetype = local\nwopiurl = http://localhost", string security = "usehttps = no", string extra = "") =>
            $"""
             [general]
             port = {FreePort()}
             loglevel = Debug
             {general}
             [security]
             wopisecretfile = {Folder.Combine("wopisecret")}
             iopsecretfile = {Folder.Combine("iopsecret")}
             {security}
             [io]
             chunksize = 4096
             recoverypath = {Folder.Combine("recovery")}
             [local]
             storagehomepath = {Folder.Combine("storage")}
             {extra}
             """;

        public static Task<WebApplication> Build(string config, bool testServer = true, StringWriter? log = null) =>
            WopiApp.Build(
                Helpers.Ini(config),
                new()
                {
                    LogWriter = log ?? new StringWriter(),
                    ConfigureBuilder = _ =>
                    {
                        if (testServer)
                        {
                            _.WebHost.UseTestServer();
                        }
                    }
                });

        public void Dispose() =>
            Folder.Dispose();
    }

    static int FreePort()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return ((IPEndPoint) listener.LocalEndpoint).Port;
    }

    // Configuration errors

    public static IEnumerable<Func<(string General, string Security, string Expected)>> InvalidConfigurations()
    {
        yield return () => ("wopiurl = http://localhost", "usehttps = no", "No option 'storagetype' in section: 'general'");
        yield return () => ("storagetype = local", "usehttps = no", "No option 'wopiurl' in section: 'general'");
        yield return () => ("storagetype = xroot\nwopiurl = http://localhost", "usehttps = no", "Storage type xroot is not supported by this port of the WOPI server");
        yield return () => ("storagetype = cs3\nwopiurl = http://localhost", "usehttps = no", "Storage type cs3 is not supported by this port of the WOPI server");
        yield return () => ("storagetype = s3\nwopiurl = http://localhost", "usehttps = no", "Unsupported/Unknown storage type s3");
        yield return () => ("storagetype = local\nwopiurl = http://localhost", "", "No option 'usehttps' in section: 'security'");
        yield return () => ("storagetype = local\nwopiurl = http://localhost\nproxiedappname = Office", "usehttps = no", "Proxed app configured but missing wopiproxy or shared key");
    }

    /// <summary>
    /// A missing or invalid option is fatal at startup.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(InvalidConfigurations))]
    public async Task InvalidConfiguration(string general, string security, string expected)
    {
        using var setup = new Setup();
        var log = new StringWriter();
        Exception? exception = null;
        WebApplication? app = null;
        try
        {
            app = await Setup.Build(setup.Config(general, security), log: log);
            // an invalid value read per request fails the request rather than the startup
            await app.StartAsync();
            using var response = await app.GetTestClient().GetAsync("/wopi/iop/test?filepath=/x&usertoken=0:0");
            await Assert.That(expected).IsEqualTo("");
        }
        catch (ConfigException configException)
        {
            exception = configException;
        }
        finally
        {
            if (app != null)
            {
                await app.DisposeAsync();
            }
        }

        if (expected.Length > 0)
        {
            await Assert.That(exception?.Message).IsEqualTo(expected);
            await Assert.That(log.ToString()).Contains("Failed to initialize the service, aborting");
        }
    }

    [Test]
    public async Task MissingSecretFile()
    {
        using var setup = new Setup();
        File.Delete(setup.Folder.Combine("iopsecret"));
        await Helpers.Throws<FileNotFoundException>(() => Setup.Build(setup.Config()));
    }

    [Test]
    public async Task MissingStorageFolder()
    {
        using var setup = new Setup();
        Directory.Delete(setup.Folder.Combine("storage"));
        var exception = await Helpers.Throws<StorageException>(() => Setup.Build(setup.Config()));
        await Assert.That(exception.Message).EndsWith(Common.EnoentMsg);
    }

    /// <summary>
    /// The secret files may end with a Windows line break.
    /// </summary>
    [Test]
    public async Task SecretWithWindowsLineBreak()
    {
        using var setup = new Setup();
        await using var app = await Setup.Build(setup.Config());
        await Assert.That(app.Services.GetRequiredService<WopiState>().WopiSecret).IsEqualTo("wopisecret");
    }

    // HTTPS

    static void WriteCertificate(string certificatePath, string keyPath)
    {
        using var key = RSA.Create(2048);
        var request = new CertificateRequest("CN=localhost", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using var certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
        File.WriteAllText(certificatePath, certificate.ExportCertificatePem());
        File.WriteAllText(keyPath, key.ExportPkcs8PrivateKeyPem());
    }

    /// <summary>
    /// With usehttps, the server listens with the PEM certificate and key of the configuration.
    /// </summary>
    [Test]
    public async Task Https()
    {
        using var setup = new Setup();
        WriteCertificate(setup.Folder.Combine("host.crt"), setup.Folder.Combine("host.key"));
        var port = FreePort();
        var config = setup.Config(
                security: $"usehttps = yes\nwopicert = {setup.Folder.Combine("host.crt")}\nwopikey = {setup.Folder.Combine("host.key")}")
            .Replace("port = ", $"port = {port}\n#");
        await using var app = await Setup.Build(config, testServer: false);
        await app.StartAsync();
        using var handler = new HttpClientHandler();
        handler.ServerCertificateCustomValidationCallback = (_, _, _, _) => true;
        using var client = new HttpClient(handler);
        using (var index = await client.GetAsync($"https://localhost:{port}/wopi"))
        {
            await Assert.That(index.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        // the WOPI validator token endpoint is for http only
        using var request = new HttpRequestMessage(HttpMethod.Get, $"https://localhost:{port}/wopi/iop/test?filepath=/x&usertoken=0:0");
        request.Headers.Add("Authorization", "Bearer iopsecret");
        using var test = await client.SendAsync(request);
        await Assert.That(test.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await app.StopAsync();
    }

    [Test]
    public async Task Https_MissingCertificate()
    {
        using var setup = new Setup();
        var log = new StringWriter();
        var config = setup.Config(security: $"usehttps = yes\nwopicert = {setup.Folder.Combine("missing.crt")}\nwopikey = {setup.Folder.Combine("missing.key")}");
        await Helpers.Throws<IOException>(() => Setup.Build(config, log: log));
        await Assert.That(log.ToString()).Contains("Failed to open the provided certificate or key to start in https mode");
    }

    // Log destinations

    [Test]
    public async Task LogWriter_File()
    {
        using var setup = new Setup();
        var path = setup.Folder.Combine("wopiserver.log");
        var config = new WopiConfig(Helpers.Ini($"[general]\nloghandler = file\nlogdest = {path}"));
        await using (var writer = WopiApp.LogWriter(config))
        {
            await writer.WriteLineAsync("line");
        }

        // appended, without a byte order mark
        await Assert.That((await File.ReadAllBytesAsync(path)).SequenceEqual(Encoding.UTF8.GetBytes($"line{Environment.NewLine}"))).IsTrue();
    }

    [Test]
    [Arguments("stdout", false)]
    [Arguments("stderr", true)]
    public async Task LogWriter_Stream(string destination, bool isError)
    {
        var config = new WopiConfig(Helpers.Ini($"[general]\nloghandler = stream\nlogdest = {destination}"));
        await Assert.That(WopiApp.LogWriter(config)).IsSameReferenceAs(isError ? Console.Error : Console.Out);
    }

    /// <summary>
    /// The JSON lines of the log, with the configured level applying to the server and framework warnings and above.
    /// </summary>
    [Test]
    public async Task Log_Format()
    {
        var writer = new StringWriter();
        using var provider = new JsonLogProvider(writer, false, "host", () => LogLevel.Information);
        var log = provider.CreateLogger("wopi");
        log.LogDebug("Not logged {key}", "value");
        log.LogWarning("Lock failed {lockop} {count} {flag} {missing}", "Lock", 3, true, null);
        log.LogError(new InvalidOperationException("boom"), "Failed");
        provider.CreateLogger("Microsoft.AspNetCore.Routing").LogInformation("Framework info");
        provider.CreateLogger("Microsoft.AspNetCore.Routing").LogWarning("Framework warning");
        var lines = writer.ToString().Trim().Split(Environment.NewLine);
        await Assert.That(lines).Count().IsEqualTo(3);
        var warning = JsonNode.Parse(lines[0])!;
        await Assert.That((string?) warning["level"]).IsEqualTo("WARNING");
        await Assert.That((string?) warning["module"]).IsEqualTo("wopi");
        await Assert.That((string?) warning["host"]).IsEqualTo("host");
        await Assert.That((string?) warning["msg"]).IsEqualTo("Lock failed");
        await Assert.That((string?) warning["count"]).IsEqualTo("3");
        await Assert.That((string?) warning["flag"]).IsEqualTo("True");
        await Assert.That((string?) warning["missing"]).IsEqualTo("None");
        await Assert.That((string?) JsonNode.Parse(lines[1])!["exception"]).Contains("boom");
        await Assert.That((string?) JsonNode.Parse(lines[2])!["msg"]).IsEqualTo("Framework warning");
    }

    // WOPI proxy

    /// <summary>
    /// A proxied app gets a WOPISrc on the WOPI proxy, whose file id is a JWT of this server's URL and the file id,
    /// signed with the proxy's shared key. Other apps get this server's WOPISrc.
    /// </summary>
    [Test]
    public async Task WopiProxy()
    {
        using var setup = new Setup();
        File.WriteAllText(setup.Folder.Combine("proxykey"), "proxykey\n");
        File.WriteAllText(setup.Folder.Combine("storage", "test.docx"), "content");
        var config = setup.Config($"storagetype = local\nwopiurl = http://localhost\nwopiproxy = https://proxy/\nwopiproxysecretfile = {setup.Folder.Combine("proxykey")}\nproxiedappname = ProxiedOffice");
        await using var app = await Setup.Build(config);
        await app.StartAsync();
        var client = app.GetTestClient();

        async Task<string> WopiSrc(string appName)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, $"/wopi/iop/openinapp?fileid=/test.docx&viewmode=VIEW_MODE_READ_WRITE&appname={appName}&appurl=https://office/edit");
            request.Headers.Add("Authorization", "Bearer iopsecret");
            request.Headers.Add("TokenHeader", "x");
            using var response = await client.SendAsync(request);
            var appUrl = (string) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["app-url"]!;
            return Uri.UnescapeDataString(appUrl[(appUrl.IndexOf("WOPISrc=", StringComparison.Ordinal) + 8)..]);
        }

        var proxied = await WopiSrc("ProxiedOffice");
        await Assert.That(proxied).StartsWith("https://proxy/wopi/files/");
        var payload = Jwt.Decode<JsonObject>(proxied["https://proxy/wopi/files/".Length..], "proxykey", 0);
        await Assert.That((string?) payload["u"]).IsEqualTo("http://localhost/wopi/files/");
        await Assert.That((string?) payload["f"]).StartsWith("local!");

        await Assert.That(await WopiSrc("Office")).StartsWith("http://localhost/wopi/files/local!");
    }

    // Command line

    [Test]
    public async Task ParseArgs()
    {
        await Assert.That(WopiProgram.ParseArgs([], _ => false)).IsEqualTo(((string?) null, "/etc/wopi/wopiserver.conf", (string?) null));
        await Assert.That(WopiProgram.ParseArgs([], _ => true)).IsEqualTo(("/etc/wopi/wopiserver.defaults.conf", "/etc/wopi/wopiserver.conf", (string?) null));
        await Assert.That(WopiProgram.ParseArgs(["--config", "a.conf", "--defaults", "b.conf"], _ => false)).IsEqualTo(("b.conf", "a.conf", (string?) null));
        await Assert.That(WopiProgram.ParseArgs(["--verbose"], _ => false).Unknown).IsEqualTo("--verbose");
        // an option without its value is unknown
        await Assert.That(WopiProgram.ParseArgs(["--config"], _ => false).Unknown).IsEqualTo("--config");
    }

    [Test]
    public async Task Run_InvalidArguments()
    {
        var error = new StringWriter();
        await Assert.That(await WopiProgram.Run(["--verbose"], error)).IsEqualTo(2);
        await Assert.That(error.ToString()).StartsWith($"Unknown argument --verbose. {WopiProgram.Usage}");
    }

    /// <summary>
    /// The shipped defaults have no storage type, so a configuration file without one fails the startup.
    /// </summary>
    [Test]
    public async Task Run_InvalidConfiguration()
    {
        using var setup = new Setup();
        var config = setup.Folder.Combine("wopiserver.conf");
        await File.WriteAllTextAsync(config, $"[general]\nloghandler = file\nlogdest = {setup.Folder.Combine("wopiserver.log")}\n");
        var error = new StringWriter();
        await Assert.That(await WopiProgram.Run(["--config", config, "--defaults", config], error)).IsEqualTo(1);
        await Assert.That(error.ToString()).StartsWith("Failed to initialize the service: No option 'storagetype' in section: 'general'");
    }
}
