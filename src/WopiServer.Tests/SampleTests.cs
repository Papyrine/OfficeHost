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
            </app>
          </net-zone>
        </wopi-discovery>
        """;

    readonly TempDirectory folder = new();
    readonly WebApplicationFactory<Program> factory;

    public SampleTests() =>
        factory = new WebApplicationFactory<Program>()
            .WithWebHostBuilder(
                builder =>
                {
                    builder.UseSetting("Sample:Files", folder.Path);
                    builder.UseSetting("Wopi:DiscoveryUrls:0", discoveryUrl);
                    builder.ConfigureTestServices(
                        services => services.AddHttpClient(WopiOptions.HttpClientName)
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
        await using var unreachable = factory.WithWebHostBuilder(_ => _.UseSetting("Wopi:DiscoveryUrls:0", "https://down/hosting/discovery"));
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
