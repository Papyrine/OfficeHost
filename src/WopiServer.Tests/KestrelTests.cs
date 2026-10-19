/// <summary>
/// The endpoints on Kestrel, for what a <see cref="TestServer"/> does not do, as limiting request bodies.
/// </summary>
public class KestrelTests
{
    static async Task<HttpResponseMessage> PutFile(TestServerHost host, int size)
    {
        await File.WriteAllTextAsync(host.PathOf("/test.docx"), "");
        var (fileId, token) = await host.Open("/test.docx");
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}/contents?access_token={token}");
        request.Content = new ByteArrayContent(new byte[size]);
        return await host.Client.SendAsync(request);
    }

    /// <summary>
    /// Office documents can be larger than the 30 MB Kestrel accepts by default: without MaxFileSize, there is no limit.
    /// </summary>
    [Test]
    public async Task PutFile_LargerThanTheServerLimit()
    {
        await using var host = await TestServerHost.Start(kestrel: true);
        using var response = await PutFile(host, 31 * 1024 * 1024);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(new FileInfo(host.PathOf("/test.docx")).Length).IsEqualTo(31 * 1024 * 1024);
    }

    [Test]
    public async Task PutFile_LargerThanMaxFileSize()
    {
        await using var host = await TestServerHost.Start(_ => _.MaxFileSize = 1000, kestrel: true);
        using var response = await PutFile(host, 1001);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.RequestEntityTooLarge);
        await Assert.That(new FileInfo(host.PathOf("/test.docx")).Length).IsEqualTo(0);
    }
}
