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
        var launch = await host.Service<WopiLauncher>().Launch(TestServerHost.Access(fileId), "edit");
        var wopiSrc = $"http://localhost/wopi/files/{fileId}";
        await Assert.That(launch.Url).IsEqualTo($"https://collabora/browser/1/cool.html?WOPISrc={Uri.EscapeDataString(wopiSrc)}");
        await Assert.That(DateTimeOffset.FromUnixTimeMilliseconds(launch.AccessTokenTtl)).IsGreaterThan(DateTimeOffset.UtcNow.AddHours(9));

        // the WOPI client calls back with the WOPISrc and the token
        using var response = await host.Client.GetAsync($"{new Uri(wopiSrc).AbsolutePath}?access_token={launch.AccessToken}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That((string?) JsonNode.Parse(await response.Content.ReadAsStringAsync())!["BaseFileName"]).IsEqualTo("test.odt");
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
        var launch = await host.Service<WopiLauncher>().Launch(TestServerHost.Access(host.Storage.GetFileId("test.docx")), action);
        await Assert.That(launch.Url).StartsWith(expected);
    }

    [Test]
    public async Task Launch_NoAction()
    {
        await using var host = await StartWithDiscovery();
        await File.WriteAllTextAsync(host.PathOf("/test.xlsx"), "content");
        var launcher = host.Service<WopiLauncher>();
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
        var launcher = host.Service<WopiLauncher>();
        var fileId = host.Storage.GetFileId("test.odt");
        var access = TestServerHost.Access(fileId);
        var exception = await Assert.ThrowsAsync<WopiClientUnreachableException>(() => launcher.Launch(access, "edit"));
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
        var launcher = host.Service<WopiLauncher>();
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
        var tokens = host.Service<IWopiTokens>();
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
        await Assert.That(await tokens.Validate(Jwt.Encode(new { fid = "file", uid = "alice", w = true, exp = 4102444800 }, "another key that is long enough, too"))).IsNull();
        time.Advance(TimeSpan.FromHours(1));
        await Assert.That(await tokens.Validate(token.Value)).IsNull();
    }

    /// <summary>
    /// The tokens are HS256 JSON Web Tokens that need an exp claim.
    /// </summary>
    [Test]
    public async Task Jwt_RoundTrip()
    {
        const string key = "a key for the JSON web tokens";
        var token = Jwt.Encode(new { a = "b", exp = 2000 }, key);
        await Assert.That(Jwt.Decode<JsonObject>(token, key, 1999)["a"]!.GetValue<string>()).IsEqualTo("b");
        Assert.Throws<JwtException>(() => Jwt.Decode<JsonObject>(token, key, 2000));
        Assert.Throws<JwtException>(() => Jwt.Decode<JsonObject>(Jwt.Encode(new { a = "b" }, key), key, 0));
        var none = $"{Base64Url(new { alg = "none" })}.{token.Split('.')[1]}.";
        Assert.Throws<JwtException>(() => Jwt.Decode<JsonObject>(none, key, 0));
    }

    static string Base64Url(object value) =>
        System.Buffers.Text.Base64Url.EncodeToString(JsonSerializer.SerializeToUtf8Bytes(value));

    [Test]
    public async Task SigningKey_Required()
    {
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => TestServerHost.Start(_ => _.SigningKey = "short"));
        await Assert.That(exception!.Message).Contains(nameof(WopiOptions.SigningKey));
    }

    /// <summary>
    /// Another <see cref="IWopiTokens"/> replaces the signed tokens, and needs no signing key.
    /// </summary>
    [Test]
    public async Task CustomTokens()
    {
        await using var host = await TestServerHost.Start(
            _ => _.SigningKey = null,
            services: _ => _.AddSingleton<IWopiTokens, OpaqueTokens>());
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "content");
        var (fileId, token) = await host.Open("/test.docx");
        await Assert.That(token).StartsWith("opaque-");
        using var response = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={token}");
        await Assert.That(await response.Content.ReadAsStringAsync()).IsEqualTo("content");
    }

    sealed class OpaqueTokens : IWopiTokens
    {
        readonly ConcurrentDictionary<string, WopiToken> tokens = new();

        public Task<WopiToken> Issue(WopiAccess access, Cancel cancel = default)
        {
            var token = new WopiToken($"opaque-{Guid.NewGuid():N}", access, DateTimeOffset.UtcNow.AddHours(1));
            tokens[token.Value] = token;
            return Task.FromResult(token);
        }

        public Task<WopiToken?> Validate(string token, Cancel cancel = default) =>
            Task.FromResult(tokens.GetValueOrDefault(token));
    }

    [Test]
    public async Task MapWopi_NeedsStorage()
    {
        var builder = WebApplication.CreateSlimBuilder();
        builder.Services.AddWopi(_ => _.SigningKey = TestServerHost.SigningKey);
        await using var app = builder.Build();
        var exception = Assert.Throws<InvalidOperationException>(() => app.MapWopi());
        await Assert.That(exception.Message).Contains(nameof(IWopiStorage));
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
    /// Without an <see cref="IWopiUserInfoStore"/>, PutUserInfo is neither advertised nor supported.
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
        await Assert.That(FileNames.SplitExtension(name)).IsEqualTo((root, extension));

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
