/// <summary>
/// The parts of the library an application uses directly: the registration, the tokens and the launcher.
/// </summary>
public class ApiTests
{
    const string discoveryUrl = "https://collabora/hosting/discovery";

    const string discovery =
        """
        <?xml version="1.0" encoding="utf-8"?>
        <wopi-discovery>
          <net-zone name="external-http">
            <app name="writer">
              <action default="true" ext="odt" name="edit" urlsrc="https://collabora/browser/1/cool.html?"/>
            </app>
            <app name="Word">
              <action name="view" ext="docx" urlsrc="https://office/wv/wordviewerframe.aspx?&lt;ui=UI_LLCC&amp;&gt;&lt;rs=DC_LLCC&amp;&gt;"/>
              <action name="edit" ext="docx" urlsrc="https://office/we/wordeditorframe.aspx?lang=en"/>
              <action name="editnew" ext="docx" requires="locks,update" urlsrc="https://office/we/wordeditorframe.aspx?new=1"/>
            </app>
          </net-zone>
        </wopi-discovery>
        """;

    static Task<HttpResponseMessage> Discovery(HttpRequestMessage request)
    {
        if (request.RequestUri!.ToString() != discoveryUrl)
        {
            throw new HttpRequestException($"Unreachable {request.RequestUri}");
        }

        return Task.FromResult(
            new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(discovery)
            });
    }

    static async Task<TestServerHost> StartWithDiscovery()
    {
        var host = await TestServerHost.Start(_ => _.DiscoveryUrls.Add(discoveryUrl));
        host.Http = Discovery;
        return host;
    }

    [Test]
    public async Task Launch()
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var fileId = host.Storage.GetFileId("test.odt");
        var launch = await host.Service<OfficeLauncher>().Launch(TestServerHost.Access(fileId), "edit");
        var wopiSrc = $"http://localhost/wopi/files/{fileId}";
        await Assert.That(launch.Url).IsEqualTo($"https://collabora/browser/1/cool.html?WOPISrc={Uri.EscapeDataString(wopiSrc)}");
        await Assert.That(DateTimeOffset.FromUnixTimeMilliseconds(launch.AccessTokenTtl)).IsGreaterThan(DateTimeOffset.UtcNow.AddHours(9));

        // the WOPI client calls back with the WOPISrc and the token
        using var response = await host.Client.GetAsync($"{new Uri(wopiSrc).AbsolutePath}?access_token={launch.AccessToken}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That((string?) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["BaseFileName"]).IsEqualTo("test.odt");
    }

    /// <summary>
    /// The host page posts the token to the editor in an iframe, with the values encoded.
    /// </summary>
    [Test]
    public async Task HostPage()
    {
        var html = new OfficeLaunch("https://editor/?a=1&b=\"2\"", "to<k>en", 42).HostPageHtml();
        await Assert.That(html).Contains("action=\"https://editor/?a=1&amp;b=&quot;2&quot;\"");
        await Assert.That(html).Contains("name=\"access_token\" value=\"to&lt;k&gt;en\"");
        await Assert.That(html).Contains("name=\"access_token_ttl\" value=\"42\"");
        await Assert.That(html).Contains("<iframe name=\"editor\"");
    }

    [Test]
    public async Task Pages()
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var fileId = Uri.EscapeDataString(host.Storage.GetFileId("test.odt"));

        using var page = await host.Client.GetAsync($"/pages/edit/{fileId}");
        await Assert.That(page.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(page.Content.Headers.ContentType!.MediaType).IsEqualTo("text/html");
        await Assert.That(await page.Content.ReadAsStringAsync()).Contains("https://collabora/browser/1/cool.html?WOPISrc=");

        // no WOPI client views .odt files
        using var noAction = await host.Client.GetAsync($"/pages/view/{fileId}");
        await Assert.That(noAction.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);

        using var missing = await host.Client.GetAsync("/pages/edit/unknown");
        await Assert.That(missing.StatusCode).IsEqualTo(HttpStatusCode.NotFound);

        using var forbidden = await host.Client.GetAsync("/pages/edit/forbidden");
        await Assert.That(forbidden.StatusCode).IsEqualTo(HttpStatusCode.Forbidden);
    }

    /// <summary>
    /// The optional placeholders of the URL templates are dropped, and WOPISrc is appended to any query.
    /// </summary>
    [Test]
    [Arguments("view", "https://office/wv/wordviewerframe.aspx?WOPISrc=")]
    [Arguments("edit", "https://office/we/wordeditorframe.aspx?lang=en&WOPISrc=")]
    public async Task Launch_UrlTemplates(string action, string expected)
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var launch = await host.Service<OfficeLauncher>().Launch(TestServerHost.Access(host.Storage.GetFileId("test.docx")), action);
        await Assert.That(launch.Url).StartsWith(expected);
    }

    /// <summary>
    /// A new document in Office for the web: the application creates an empty file in its storage and launches the
    /// editnew action on it; the client's first save is a PutFile without a lock, which an empty file accepts
    /// ([MS-WOPI] 3.3.5.3.2), and after it the file is no longer empty.
    /// </summary>
    [Test]
    public async Task Launch_EditNew()
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllBytesAsync(host.PathOf("/new.docx"), []);
        var fileId = host.Storage.GetFileId("new.docx");
        var launch = await host.Service<OfficeLauncher>().Launch(TestServerHost.Access(fileId), "editnew");
        await Assert.That(launch.Url).StartsWith("https://office/we/wordeditorframe.aspx?new=1&WOPISrc=");

        var path = $"/wopi/files/{Uri.EscapeDataString(fileId)}";
        using var info = await host.Client.GetAsync($"{path}?access_token={launch.AccessToken}");
        await Assert.That((long?) JsonNode.Parse(await info.Content.ReadAsStringAsync())!["Size"]).IsEqualTo(0);

        static HttpRequestMessage PutFile(string path, string token, string content)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, $"{path}/contents?access_token={token}")
            {
                Content = new StringContent(content)
            };
            request.Headers.Add("X-WOPI-Override", "PUT");
            return request;
        }

        using var first = await host.Client.SendAsync(PutFile(path, launch.AccessToken, "new content"));
        await Assert.That(first.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/new.docx"))).IsEqualTo("new content");

        using var second = await host.Client.SendAsync(PutFile(path, launch.AccessToken, "other"));
        await Assert.That(second.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
    }

    [Test]
    public async Task Launch_NoAction()
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllTextAsync(host.PathOf("/test.xlsx"), "content");
        var launcher = host.Service<OfficeLauncher>();
        var fileId = host.Storage.GetFileId("test.xlsx");
        var access = TestServerHost.Access(fileId);
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => launcher.Launch(access, "edit"));
        await Assert.That(exception!.Message).Contains(".xlsx");
    }

    [Test]
    public async Task Launch_DiscoveryUnreachable()
    {
        await using var host = await TestServerHost.Start(_ => _.DiscoveryUrls.Add(discoveryUrl));
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var launcher = host.Service<OfficeLauncher>();
        var fileId = host.Storage.GetFileId("test.odt");
        var access = TestServerHost.Access(fileId);
        var exception = await Assert.ThrowsAsync<OfficeClientUnreachableException>(() => launcher.Launch(access, "edit"));
        await Assert.That(exception!.Message).IsEqualTo($"No WOPI client offers the edit action for .odt files, and the discovery of {discoveryUrl} could not be fetched");
        await Assert.That(exception.DiscoveryUrls).IsEquivalentTo([discoveryUrl]);
    }

    /// <summary>
    /// Without BaseUrl, WOPISrc is built from the request the application launches from.
    /// </summary>
    [Test]
    public async Task Launch_FromRequest()
    {
        await using var host = await TestServerHost.Start(
            _ =>
            {
                _.BaseUrl = null;
                _.DiscoveryUrls.Add(discoveryUrl);
            });
        host.Http = Discovery;
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var access = TestServerHost.Access(host.Storage.GetFileId("test.odt"));
        var launcher = host.Service<OfficeLauncher>();
        await Assert.ThrowsAsync<InvalidOperationException>(() => launcher.Launch(access, "edit"));
        var context = new DefaultHttpContext
        {
            Request =
            {
                Scheme = "https",
                Host = new("app.example"),
                PathBase = "/docs"
            }
        };
        var launch = await launcher.Launch(access, "edit", context.Request);
        await Assert.That(launch.Url).Contains(Uri.EscapeDataString($"https://app.example/docs/wopi/files/{access.FileId}"));
    }

    /// <summary>
    /// A token grants its access to its file only, until it expires.
    /// </summary>
    [Test]
    public async Task Tokens()
    {
        var time = new ManualTime();
        await using var host = await TestServerHost.Start(_ => _.TokenLifetime = TimeSpan.FromHours(1), time);
        var tokens = host.Service<IOfficeTokens>();
        var access = TestServerHost.Access("file", Mode.ViewOnly) with
        {
            UserName = "Bob Example"
        };
        var token = await tokens.Issue(access);
        await Assert.That(token.Expires).IsEqualTo(DateTimeOffset.FromUnixTimeSeconds((time.Now + TimeSpan.FromHours(1)).ToUnixTimeSeconds()));
        var validated = (await tokens.Validate(token.Value))!;
        await Assert.That(validated.Access).IsEqualTo(access);
        await Assert.That(validated.Expires).IsEqualTo(token.Expires);

        await Assert.That(await tokens.Validate($"{token.Value}x")).IsNull();
        await Assert.That(await tokens.Validate("not a token")).IsNull();
        var otherKey = "another key that is long enough, too"u8.ToArray();
        await Assert.That(await tokens.Validate(SignedToken.Encode(TestServerHost.Access("file"), DateTimeOffset.UtcNow.AddHours(1), otherKey))).IsNull();
        time.Advance(TimeSpan.FromHours(1));
        await Assert.That(await tokens.Validate(token.Value)).IsNull();
    }

    /// <summary>
    /// The signed tokens hold every right of the access, and nothing but a token signed with the key, whole and unchanged,
    /// is one.
    /// </summary>
    [Test]
    public async Task SignedToken_RoundTrip()
    {
        var key = "a key for the signed tokens, long enough"u8.ToArray();
        var expires = DateTimeOffset.FromUnixTimeSeconds(2000);
        var now = DateTimeOffset.FromUnixTimeSeconds(1999);
        OfficeAccess[] accesses =
        [
            new("file", "user"),
            new("fïlé/ünicode", "üser")
            {
                UserName = "Zoë Example",
                IsAnonymous = true,
                CanWrite = true,
                CanWriteRelative = true,
                CanRename = true,
                ViewOnly = true
            },
            new("file", "user")
            {
                CanWriteRelative = true
            }
        ];
        foreach (var access in accesses)
        {
            var token = SignedToken.Encode(access, expires, key);
            await Assert.That(SignedToken.Decode(token, key, now)).IsEqualTo((access, expires));
        }

        var valid = SignedToken.Encode(accesses[1], expires, key);
        await Assert.That(SignedToken.Decode(valid, key, expires)).IsNull();
        await Assert.That(SignedToken.Decode(valid[..^1], key, now)).IsNull();
        await Assert.That(SignedToken.Decode(valid[1..], key, now)).IsNull();
        await Assert.That(SignedToken.Decode("", key, now)).IsNull();
        await Assert.That(SignedToken.Decode("not a token!", key, now)).IsNull();
        await Assert.That(SignedToken.Decode(new('A', 5000), key, now)).IsNull();

        // any byte changed fails the signature
        var bytes = System.Buffers.Text.Base64Url.DecodeFromChars(valid);
        for (var index = 0; index < bytes.Length; index++)
        {
            var changed = bytes.ToArray();
            changed[index] ^= 0x01;
            await Assert.That(SignedToken.Decode(System.Buffers.Text.Base64Url.EncodeToString(changed), key, now)).IsNull();
        }
    }

    /// <summary>
    /// A token for an access of the sample is short, so that a desktop site URL with a file name stays under the 255
    /// characters Word accepts for a Save As.
    /// </summary>
    [Test]
    public async Task SignedToken_Short()
    {
        var access = new OfficeAccess("jCT-Sx_4ei7yZbxe", "demo")
        {
            UserName = "Demo User",
            CanWrite = true,
            CanWriteRelative = true,
            CanRename = true
        };
        var token = SignedToken.Encode(access, DateTimeOffset.UtcNow.AddHours(10), "a key for the signed tokens, long enough"u8.ToArray());
        await Assert.That(token.Length).IsLessThan(80);
    }

    [Test]
    public async Task SigningKey_Required()
    {
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => TestServerHost.Start(_ => _.SigningKey = "short"));
        await Assert.That(exception!.Message).Contains(nameof(OfficeHostOptions.SigningKey));
    }

    /// <summary>
    /// Another <see cref="IOfficeTokens"/> replaces the signed tokens, and needs no signing key.
    /// </summary>
    [Test]
    public async Task CustomTokens()
    {
        await using var host = await TestServerHost.Start(
            _ => _.SigningKey = null,
            services: _ => _.AddSingleton<IOfficeTokens, OpaqueTokens>());
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token) = await host.Open("/test.docx");
        await Assert.That(token).StartsWith("opaque-");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("content");
    }

    sealed class OpaqueTokens : IOfficeTokens
    {
        ConcurrentDictionary<string, OfficeToken> tokens = new();

        public Task<OfficeToken> Issue(OfficeAccess access, Cancel cancel = default)
        {
            var token = new OfficeToken($"opaque-{Guid.NewGuid():N}", access, DateTimeOffset.UtcNow.AddHours(1));
            tokens[token.Value] = token;
            return Task.FromResult(token);
        }

        public Task<OfficeToken?> Validate(string token, Cancel cancel = default) =>
            Task.FromResult(tokens.GetValueOrDefault(token));
    }

    /// <summary>
    /// Without an <see cref="IOfficeStorage"/>, MapOfficeHost says so. In Production: in Development, Build validates the
    /// services and fails first, whatever the environment of the test runner.
    /// </summary>
    [Test]
    public async Task MapOfficeHost_NeedsStorage()
    {
        var builder = WebApplication.CreateSlimBuilder(
            new WebApplicationOptions
            {
                EnvironmentName = Environments.Production
            });
        builder.Services.AddOfficeHost(_ => _.SigningKey = TestServerHost.SigningKey);
        await using var app = builder.Build();
        var exception = Assert.Throws<InvalidOperationException>(() => app.MapOfficeHost());
        await Assert.That(exception.Message).Contains(nameof(IOfficeStorage));
    }

    /// <summary>
    /// A scoped storage and scoped tokens, as over a database context: each request has its own.
    /// </summary>
    [Test]
    public async Task ScopedServices()
    {
        using var folder = new TempDirectory();
        await File.WriteAllTextAsync(Path.Combine(folder, "test.docx"), "content");
        var storage = new FileSystemStorage(folder);
        var builder = WebApplication.CreateSlimBuilder();
        builder.WebHost.UseTestServer();
        builder.Host.UseDefaultServiceProvider(
            _ =>
            {
                _.ValidateScopes = true;
                _.ValidateOnBuild = true;
            });
        var storages = 0;
        var services = builder.Services;
        services.AddScoped<IOfficeStorage>(
            _ =>
            {
                Interlocked.Increment(ref storages);
                return new ForwardingStorage(storage);
            });
        services.AddSingleton<OpaqueTokens>();
        services.AddScoped<IOfficeTokens>(_ => _.GetRequiredService<OpaqueTokens>());
        services.AddOfficeHost(_ => { });
        await using var app = builder.Build();
        app.MapOfficeHost();
        await app.StartAsync();
        string token;
        using (var scope = app.Services.CreateScope())
        {
            scope.ServiceProvider.GetRequiredService<OfficeLauncher>();
            token = (await scope.ServiceProvider.GetRequiredService<IOfficeTokens>().Issue(TestServerHost.Access(storage.GetFileId("test.docx")))).Value;
        }

        var before = storages;
        using var client = app.GetTestClient();
        var fileId = storage.GetFileId("test.docx");
        using (var info = await client.GetAsync($"/wopi/files/{fileId}?access_token={token}"))
        {
            await Assert.That(info.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var content = await client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}"))
        {
            await Assert.That(await content.Content.ReadAsStringAsync()).IsEqualTo("content");
        }

        await Assert.That(storages).IsEqualTo(before + 2);
    }

    /// <summary>
    /// A WOPI client that is slow to answer delays only the requests that need its discovery.
    /// </summary>
    [Test]
    public async Task Discovery_SlowClient()
    {
        const string slowUrl = "https://slow/hosting/discovery";
        var slow = new TaskCompletionSource<HttpResponseMessage>(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var host = await TestServerHost.Start(_ => _.DiscoveryUrls.Add(discoveryUrl));
        host.Http = request =>
        {
            if (request.RequestUri!.ToString() == slowUrl)
            {
                return slow.Task;
            }

            return Discovery(request);
        };
        await File.WriteAllTextAsync(host.PathOf("/test.odt"), "content");
        var access = TestServerHost.Access(host.Storage.GetFileId("test.odt"));
        var launcher = host.Service<OfficeLauncher>();
        await launcher.Launch(access, "edit");
        var pending = launcher.Launch(access, "edit", discoveryUrl: slowUrl);
        var launch = await launcher.Launch(access, "edit").WaitAsync(TimeSpan.FromSeconds(10));
        await Assert.That(launch.Url).StartsWith("https://collabora/");
        await Assert.That(pending.IsCompleted).IsFalse();
        slow.SetResult(
            new(HttpStatusCode.OK)
            {
                Content = new StringContent(discovery)
            });
        await Assert.That((await pending).Url).StartsWith("https://collabora/");
    }

    /// <summary>
    /// The application adds properties to CheckFileInfo, and gives the host pages of the files.
    /// </summary>
    [Test]
    public async Task CheckFileInfo_Customized()
    {
        await using var host = await TestServerHost.Start(
            _ => _.OnCheckFileInfo = context =>
            {
                context.Properties["CloseUrl"] = $"https://app/folder?file={context.File.Name}";
                context.Properties["UserCanWrite"] = false;
                return Task.CompletedTask;
            });
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token) = await host.Open("/test.docx");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}");
        var info = JsonNode.Parse(await response.Content.ReadAsStringAsync())!;
        await Assert.That((string?) info["CloseUrl"]).IsEqualTo("https://app/folder?file=test.docx");
        await Assert.That((bool?) info["UserCanWrite"]).IsFalse();
        await Assert.That((string?) info["HostEditUrl"]).IsEqualTo($"https://app/edit/{fileId}");
        await Assert.That((string?) info["HostViewUrl"]).IsEqualTo($"https://app/view/{fileId}");
    }

    /// <summary>
    /// Without an <see cref="IOfficeUserInfoStore"/>, PutUserInfo is neither advertised nor supported.
    /// </summary>
    [Test]
    public async Task UserInfo_Unsupported()
    {
        await using var host = await TestServerHost.Start(userInfo: false);
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token) = await host.Open("/test.docx");
        using (var info = await host.Client.GetAsync($"/wopi/files/{fileId}?access_token={token}"))
        {
            await Assert.That((bool?) JsonNode.Parse(await info.Content.ReadAsStringAsync())!["SupportsUserInfo"]).IsFalse();
        }

        using var response = await host.Post(fileId, token, "PUT_USER_INFO");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
    }

    [Test]
    [Arguments("report.docx", "report", ".docx")]
    [Arguments("archive.tar.gz", "archive.tar", ".gz")]
    [Arguments("noextension", "noextension", "")]
    [Arguments(".profile", ".profile", "")]
    [Arguments("..double", "..double", "")]
    public async Task FileNames_SplitExtension(string name, string root, string extension) =>
        await Assert.That(Split(name)).IsEqualTo((root, extension));

    static (string Root, string Extension) Split(string name)
    {
        var (root, extension) = FileNames.SplitExtension(name);
        return (root.ToString(), extension.ToString());
    }

    [Test]
    public async Task FileNames_Utf7()
    {
        await Assert.That(FileNames.DecodeUtf7("+AOk-t+AOk-.docx")).IsEqualTo("été.docx");
        await Assert.That(FileNames.EncodeUtf7("été.docx")).IsEqualTo("+AOk-t+AOk-.docx");
    }

    [Test]
    public async Task FileNames_Numbered()
    {
        await Assert.That(FileNames.Numbered("report.docx", 2)).IsEqualTo("report (2).docx");
        var numbered = FileNames.Numbered($"{new string('x', 300)}.docx", 12);
        await Assert.That(numbered).EndsWith("x (12).docx");
        await Assert.That(Encoding.UTF8.GetByteCount(numbered)).IsEqualTo(FileNames.MaxBytes);
    }
}
