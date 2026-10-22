using Microsoft.AspNetCore.Mvc.Testing;

/// <summary>
/// The sample application, with a stubbed WOPI client discovery: the file list, the pages that launch the editor, and
/// the WOPI requests of the editor with the token they give.
/// </summary>
public partial class SampleTests : IAsyncDisposable
{
    const string discoveryUrl = "https://collabora/hosting/discovery";

    const string discovery =
        """
        <wopi-discovery>
          <net-zone name="external-http">
            <app name="writer">
              <action name="edit" ext="docx" urlsrc="https://collabora/browser/1/cool.html?"/>
              <action name="view" ext="docx" urlsrc="https://collabora/browser/1/cool.html?permission=readonly&amp;"/>
              <action name="editnew" ext="docx" urlsrc="https://collabora/browser/1/cool.html?new=docx&amp;"/>
            </app>
            <app name="calc">
              <action name="editnew" ext="xlsx" urlsrc="https://collabora/browser/1/cool.html?new=xlsx&amp;"/>
            </app>
          </net-zone>
        </wopi-discovery>
        """;

    TempDirectory folder = new();
    WebApplicationFactory<Program> factory;

    public SampleTests() =>
        factory = new WebApplicationFactory<Program>()
            .WithWebHostBuilder(
                builder =>
                {
                    builder.UseSetting("Sample:Files", folder.Path);
                    builder.UseSetting("OfficeHost:DiscoveryUrls:0", discoveryUrl);
                    builder.ConfigureTestServices(
                        services => services.AddHttpClient(OfficeHostOptions.HttpClientName)
                            .ConfigurePrimaryHttpMessageHandler(() => new DiscoveryHandler()));
                });

    public async ValueTask DisposeAsync()
    {
        await factory.DisposeAsync();
        folder.Dispose();
    }

    sealed class DiscoveryHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, Cancel cancellationToken)
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
    }

    [GeneratedRegex("""href="/edit/([^"]+)">""")]
    private static partial Regex EditLink();

    [GeneratedRegex("""action="([^"]+)"[\s\S]*name="access_token" value="([^"]+)"[\s\S]*name="access_token_ttl" value="(\d+)">""")]
    private static partial Regex LaunchForm();

    static async Task<string> WelcomeId(HttpClient client)
    {
        var files = await client.GetStringAsync("/");
        await Assert.That(files).Contains("Welcome.docx");
        return Uri.UnescapeDataString(EditLink().Match(files).Groups[1].Value);
    }

    /// <summary>
    /// The page of an action posts the token to the URL of the editor, which then calls back the WOPI endpoints at
    /// BaseUrl with that token.
    /// </summary>
    [Test]
    [Arguments("edit", "https://collabora/browser/1/cool.html?WOPISrc=", true)]
    [Arguments("view", "https://collabora/browser/1/cool.html?permission=readonly&WOPISrc=", false)]
    public async Task Launch(string action, string editorUrl, bool canWrite)
    {
        using var client = factory.CreateClient();
        var fileId = await WelcomeId(client);
        var page = await client.GetStringAsync($"/{action}/{Uri.EscapeDataString(fileId)}");
        var form = LaunchForm().Match(page);
        await Assert.That(form.Success).IsTrue();
        var url = WebUtility.HtmlDecode(form.Groups[1].Value);
        var wopiSrc = $"http://host.docker.internal:5000/wopi/files/{fileId}";
        await Assert.That(url).IsEqualTo($"{editorUrl}{Uri.EscapeDataString(wopiSrc)}");

        var token = WebUtility.HtmlDecode(form.Groups[2].Value);
        using var response = await client.GetAsync($"/wopi/files/{fileId}?access_token={Uri.EscapeDataString(token)}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var info = JsonNode.Parse(await response.Content.ReadAsStringAsync())!;
        await Assert.That((string?) info["BaseFileName"]).IsEqualTo("Welcome.docx");
        await Assert.That((bool?) info["UserCanWrite"]).IsEqualTo(canWrite);
        await Assert.That((string?) info["UserFriendlyName"]).IsEqualTo("Demo User");
        await Assert.That((string?) info["CloseUrl"]).IsEqualTo("http://localhost:5000");
        await Assert.That((string?) info["HostEditUrl"]).IsEqualTo($"http://localhost:5000/edit/{Uri.EscapeDataString(fileId)}");
    }

    /// <summary>
    /// Sample:RequestLog logs every request and response to a folder, after the files already there, with the token of
    /// a desktop site out of the file names.
    /// </summary>
    [Test]
    public async Task RequestLog()
    {
        using var logs = new TempDirectory();
        await File.WriteAllTextAsync(Path.Combine(logs.Path, "0001-GET-earlier.txt"), "");
        await using var logged = factory.WithWebHostBuilder(_ => _.UseSetting("Sample:RequestLog", logs.Path));
        using var client = logged.CreateClient();
        var fileId = await WelcomeId(client);
        var token = (await logged.Services.GetRequiredService<IOfficeTokens>().Issue(new(fileId, "demo"))).Value;
        using var options = await client.SendAsync(new(HttpMethod.Options, $"/fsshttp/{token}/Welcome.docx"));

        var names = Directory.EnumerateFiles(logs.Path).Select(_ => Path.GetFileName(_)).Order().ToList();
        await Assert.That(names).IsEquivalentTo(["0001-GET-earlier.txt", "0002-GET-.txt", "0003-OPTIONS-fsshttp_~_Welcome.docx.txt"]);
        var files = await File.ReadAllTextAsync(Path.Combine(logs.Path, "0002-GET-.txt"));
        await Assert.That(files).Contains("--- response 200");
        await Assert.That(files).Contains("Welcome.docx");
        var discovery = await File.ReadAllTextAsync(Path.Combine(logs.Path, "0003-OPTIONS-fsshttp_~_Welcome.docx.txt"));
        await Assert.That(discovery).Contains("X-MSFSSHTTP: 1.3");
    }

    [GeneratedRegex("""id="open" href="([^"]+)">""")]
    private static partial Regex DesktopLink();

    /// <summary>
    /// The desktop page links to the file in Word, as the user of ?user=, so that two people can coauthor the demo.
    /// </summary>
    [Test]
    [Arguments("", "demo", "Demo User")]
    [Arguments("?user=bob", "bob", "Bob User")]
    public async Task Desktop(string query, string userId, string userName)
    {
        using var client = factory.CreateClient();
        var fileId = await WelcomeId(client);
        var page = await client.GetStringAsync($"/desktop/{Uri.EscapeDataString(fileId)}{query}");
        var link = WebUtility.HtmlDecode(DesktopLink().Match(page).Groups[1].Value);
        await Assert.That(link).StartsWith("ms-word:ofe|u|http://localhost/fsshttp/");
        await Assert.That(link).EndsWith("/Welcome.docx");
        var token = link.Split('/')[^2];
        var access = (await factory.Services.GetRequiredService<IOfficeTokens>().Validate(token))!.Access;
        await Assert.That(access.UserId).IsEqualTo(userId);
        await Assert.That(access.UserName).IsEqualTo(userName);
    }

    /// <summary>
    /// A new document is an empty file of the first free name, opened with the editnew action of the WOPI client,
    /// whose first save writes it, or in desktop Office.
    /// </summary>
    [Test]
    public async Task NewDocument_Browser()
    {
        using var client = factory.CreateClient(new() { AllowAutoRedirect = false });
        var location = await New(client, "xlsx", "browser");
        var fileId = factory.Services.GetRequiredService<FileSystemStorage>().GetFileId("Workbook 1.xlsx");
        await Assert.That(location).IsEqualTo($"/editnew/{Uri.EscapeDataString(fileId)}");
        await Assert.That(new FileInfo(Path.Combine(folder, "Workbook 1.xlsx")).Length).IsEqualTo(0);
        await Assert.That(await New(client, "xlsx", "browser")).Contains("/editnew/");
        await Assert.That(File.Exists(Path.Combine(folder, "Workbook 2.xlsx"))).IsTrue();

        var page = await client.GetStringAsync(location);
        var form = LaunchForm().Match(page);
        await Assert.That(WebUtility.HtmlDecode(form.Groups[1].Value)).StartsWith("https://collabora/browser/1/cool.html?new=xlsx&WOPISrc=");

        // the first save of the editor needs no lock, as the file is empty
        var token = WebUtility.HtmlDecode(form.Groups[2].Value);
        using var save = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{Uri.EscapeDataString(fileId)}/contents?access_token={Uri.EscapeDataString(token)}")
        {
            Content = new ByteArrayContent("saved"u8.ToArray())
        };
        save.Headers.Add("X-WOPI-Override", "PUT");
        using var saved = await client.SendAsync(save);
        await Assert.That(saved.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await File.ReadAllTextAsync(Path.Combine(folder, "Workbook 1.xlsx"))).IsEqualTo("saved");
    }

    [Test]
    public async Task NewDocument_Desktop()
    {
        using var client = factory.CreateClient(new() { AllowAutoRedirect = false });
        var location = await New(client, "docx", "desktop");
        var fileId = factory.Services.GetRequiredService<FileSystemStorage>().GetFileId("Document 1.docx");
        await Assert.That(location).IsEqualTo($"/desktop/{Uri.EscapeDataString(fileId)}");
        var page = await client.GetStringAsync(location);
        var link = WebUtility.HtmlDecode(DesktopLink().Match(page).Groups[1].Value);
        await Assert.That(link).StartsWith("ms-word:ofe|u|http://localhost/fsshttp/");
        await Assert.That(link).EndsWith("/Document%201.docx");
    }

    static async Task<string> New(HttpClient client, string type, string open)
    {
        using var response = await client.PostAsync(
            "/new",
            new FormUrlEncodedContent([new("type", type), new("open", open)]));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Redirect);
        return response.Headers.Location!.ToString();
    }

    [Test]
    public async Task Launch_UnsupportedFile()
    {
        await File.WriteAllTextAsync(Path.Combine(folder, "notes.txt"), "text");
        using var client = factory.CreateClient();
        var fileId = factory.Services.GetRequiredService<FileSystemStorage>().GetFileId("notes.txt");
        using var response = await client.GetAsync($"/edit/{Uri.EscapeDataString(fileId)}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(await response.Content.ReadAsStringAsync()).Contains("No WOPI client offers the edit action for .txt files");
    }

    /// <summary>
    /// When the WOPI client is down, the page says how to start Collabora Online.
    /// </summary>
    [Test]
    public async Task Launch_ClientUnreachable()
    {
        await using var unreachable = factory.WithWebHostBuilder(_ => _.UseSetting("OfficeHost:DiscoveryUrls:0", "https://down/hosting/discovery"));
        using var client = unreachable.CreateClient();
        var fileId = await WelcomeId(client);
        using var response = await client.GetAsync($"/edit/{Uri.EscapeDataString(fileId)}");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.ServiceUnavailable);
        var page = WebUtility.HtmlDecode(await response.Content.ReadAsStringAsync());
        await Assert.That(page).Contains("the discovery of https://down/hosting/discovery could not be fetched");
        await Assert.That(page).Contains("Is Collabora Online running?");
        await Assert.That(page).Contains("docker run --rm --publish 9980:9980");
    }

    [Test]
    public async Task Launch_UnknownFile()
    {
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/edit/unknown");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }
}
