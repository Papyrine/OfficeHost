using System.Net.Http.Headers;
using System.Text;
using System.Xml.Linq;
using Microsoft.AspNetCore.Mvc.Testing;

/// <summary>
/// The throwaway FSSHTTP spike of the sample (docs/fsshttp-spike.md): the SharePoint-like discovery answers, and the
/// cell storage service answering the request of the example of [MS-FSSHTTP] 4.1.1.
/// </summary>
public class FsshttpSpikeTests : IAsyncDisposable
{
    static readonly XNamespace soap = "http://schemas.microsoft.com/sharepoint/soap/";

    TempDirectory folder = new();
    TempDirectory logs = new();
    WebApplicationFactory<Program> factory;

    public FsshttpSpikeTests() =>
        factory = new WebApplicationFactory<Program>()
            .WithWebHostBuilder(
                builder =>
                {
                    builder.UseSetting("Sample:Files", folder.Path);
                    builder.UseSetting("Spike:Fsshttp", "true");
                    builder.UseSetting("Spike:Logs", logs.Path);
                });

    public async ValueTask DisposeAsync()
    {
        await factory.DisposeAsync();
        folder.Dispose();
        logs.Dispose();
    }

    async Task<string> Token()
    {
        var storage = factory.Services.GetRequiredService<FileSystemStorage>();
        var tokens = factory.Services.GetRequiredService<IOfficeTokens>();
        var token = await tokens.Issue(new(storage.GetFileId("Welcome.docx"), "demo") { UserName = "Demo User" });
        return token.Value;
    }

    [Test]
    public async Task Options_AnswersAsSharePoint()
    {
        using var client = factory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Options, "/");
        using var response = await client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(response.Headers.GetValues("MicrosoftSharePointTeamServices").Single()).IsEqualTo("16.0.0.0");
        await Assert.That(response.Headers.GetValues("MS-Author-Via").Single()).IsEqualTo("MS-FP/4.0,DAV");
        await Assert.That(response.Headers.GetValues("X-MSFSSHTTP").Single()).IsEqualTo("1.3");
        await Assert.That(response.Headers.Contains("X-MSWEBURL")).IsFalse();
        await Assert.That(Directory.EnumerateFiles(logs.Path).Count()).IsEqualTo(1);
    }

    /// <summary>
    /// Word asks for the site of a document with X-MSGETWEBURL, and SharePoint answers with its path in X-MSWEBURL:
    /// with X-MSFSSHTTP, that is what makes Word use cellstorage.svc (docs/fsshttp-spike.md).
    /// </summary>
    [Test]
    public async Task Options_WebUrl()
    {
        var token = await Token();
        using var client = factory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Options, $"/t/{token}/");
        request.Headers.Add("X-MSGETWEBURL", "t");
        using var response = await client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(response.Headers.GetValues("X-MSWEBURL").Single()).IsEqualTo($"/t/{token}");
    }

    [Test]
    public async Task Basic_Challenges()
    {
        using var client = factory.CreateClient();
        using var response = await client.GetAsync("/b/Welcome.docx");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        await Assert.That(response.Headers.WwwAuthenticate.Single().Scheme).IsEqualTo("Basic");
    }

    [Test]
    public async Task Basic_TokenAsPassword()
    {
        using var client = factory.CreateClient();
        var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"someone:{await Token()}"));
        client.DefaultRequestHeaders.Authorization = new("Basic", credentials);
        using var response = await client.GetAsync("/b/Welcome.docx");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    [Test]
    public async Task TokenInPath_GetsDocument()
    {
        using var client = factory.CreateClient();
        using var response = await client.GetAsync($"/t/{await Token()}/Welcome.docx");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(response.Content.Headers.ContentType!.MediaType)
            .IsEqualTo("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    }

    /// <summary>
    /// The request of [MS-FSSHTTP] 4.1.1 in MTOM: WhoAmI, ServerTime and the Coauth join succeed, the SchemaLock
    /// depending on the join is not run, and the Cell subrequests are not supported by the spike.
    /// </summary>
    [Test]
    public async Task CellStorage_SpecExample()
    {
        var token = await Token();
        using var client = factory.CreateClient();
        var url = $"http://localhost/t/{token}/Welcome.docx";
        var xml =
            $"""
             <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
               <s:Body>
                 <RequestVersion Version="2" MinorVersion="2" xmlns="http://schemas.microsoft.com/sharepoint/soap/"/>
                 <RequestCollection CorrelationId="{Guid.NewGuid()}" xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                   <Request Url="{url}" RequestToken="1">
                     <SubRequest Type="Coauth" SubRequestToken="1">
                       <SubRequestData CoauthRequestType="JoinCoauthoring" SchemaLockID="29358EC1-E813-4793-8E70-ED0344E7B73C" ClientID="{Guid.NewGuid()}" Timeout="3600"/>
                     </SubRequest>
                     <SubRequest Type="SchemaLock" SubRequestToken="2" DependsOn="1" DependencyType="OnNotSupported">
                       <SubRequestData SchemaLockRequestType="GetLock" SchemaLockID="29358EC1-E813-4793-8E70-ED0344E7B73C" ClientID="{Guid.NewGuid()}" Timeout="3600"/>
                     </SubRequest>
                     <SubRequest Type="Cell" SubRequestToken="3" DependsOn="2" DependencyType="OnExecute">
                       <SubRequestData PartitionID="383adc0b-e66e-4438-95e6-e39ef9720122" BinaryDataSize="88">DAALAJzPKfM5lAabBgIAAO4CAACqAiAAfrgx50XdqkSrgAx1+9FTDnoCCAD0J0UfdwEWAgYAAwUAigICAADaAgYAAwAAygIIAAgAgAOEAEELAawCAFUDAQ==</SubRequestData>
                     </SubRequest>
                     <SubRequest Type="ServerTime" SubRequestToken="5"/>
                     <SubRequest Type="WhoAmI" SubRequestToken="7"/>
                     <SubRequest Type="GetDocMetaInfo" SubRequestToken="8"/>
                   </Request>
                 </RequestCollection>
               </s:Body>
             </s:Envelope>
             """;
        using var content = new MultipartContent("related", "boundary-of-the-test");
        content.Headers.ContentType!.Parameters.Add(new("type", "\"application/xop+xml\""));
        var root = new StringContent(xml, Encoding.UTF8);
        root.Headers.ContentType = MediaTypeHeaderValue.Parse("application/xop+xml;charset=utf-8;type=\"text/xml\"");
        content.Add(root);
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/t/{token}/_vti_bin/cellstorage.svc/CellStorageService")
        {
            Content = content
        };
        request.Headers.Add("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest");
        using var response = await client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(response.Content.Headers.ContentType!.MediaType).IsEqualTo("multipart/related");

        var body = await response.Content.ReadAsStringAsync();
        var start = body.IndexOf("<s:Envelope", StringComparison.Ordinal);
        var end = body.IndexOf("</s:Envelope>", StringComparison.Ordinal) + "</s:Envelope>".Length;
        var envelope = XDocument.Parse(body[start..end]);
        var collection = envelope.Descendants(soap + "ResponseCollection").Single();
        await Assert.That((string?) collection.Attribute("WebUrl")).IsEqualTo($"http://localhost/t/{token}");
        var codes = envelope.Descendants(soap + "SubResponse")
            .ToDictionary(_ => (string) _.Attribute("SubRequestToken")!, _ => (string) _.Attribute("ErrorCode")!);
        await Assert.That(codes["1"]).IsEqualTo("Success");
        await Assert.That(codes["2"]).IsEqualTo("DependentOnlyOnNotSupportedRequestGetSupported");
        await Assert.That(codes["3"]).IsEqualTo("RequestNotSupported");
        await Assert.That(codes["5"]).IsEqualTo("Success");
        await Assert.That(codes["7"]).IsEqualTo("Success");
        await Assert.That(codes["8"]).IsEqualTo("Success");

        var whoAmI = envelope.Descendants(soap + "SubResponse").Single(_ => (string?) _.Attribute("SubRequestToken") == "7");
        await Assert.That((string?) whoAmI.Element(soap + "SubResponseData")!.Attribute("UserName")).IsEqualTo("Demo User");
    }
}
