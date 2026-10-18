using Microsoft.AspNetCore.WebUtilities;

/// <summary>
/// The WOPI bridge with a stubbed CodiMD.
/// </summary>
public class BridgeTests
{
    /// <summary>
    /// The CodiMD endpoints used by the bridge, holding the notes in memory.
    /// </summary>
    sealed class StubCodiMD
    {
        public Dictionary<string, string> Notes { get; } = [];

        public async Task<HttpResponseMessage> Handle(HttpRequestMessage request)
        {
            var path = request.RequestUri!.AbsolutePath;
            if (request.Method == HttpMethod.Head)
            {
                return new(HttpStatusCode.OK);
            }

            if (request.Method == HttpMethod.Put &&
                path.StartsWith("/api/notes/", StringComparison.Ordinal))
            {
                var json = JsonNode.Parse(await request.Content!.ReadAsStringAsync())!;
                Notes[path["/api/notes/".Length..]] = (string) json["content"]!;
                return new(HttpStatusCode.OK);
            }

            if (request.Method == HttpMethod.Get &&
                path.EndsWith("/download", StringComparison.Ordinal))
            {
                return new(HttpStatusCode.OK)
                {
                    Content = new StringContent(Notes[path[1..^"/download".Length]])
                };
            }

            return new(HttpStatusCode.NotFound);
        }
    }

    [Test]
    public async Task CodiMD_OpenEditSave()
    {
        var codimd = new StubCodiMD();
        await using var host = await TestServerHost.Start();
        host.Apps = codimd.Handle;
        await File.WriteAllTextAsync(host.PathOf("/notes.md"), "# Hello\r\n");

        // open: the note is pushed to CodiMD, with Windows line endings converted
        using var open = await host.OpenInApp("/notes.md", appName: "CodiMD", appUrl: "http://codimd");
        await Assert.That(open.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var appUrl = (string) JsonNode.Parse(await open.Content.ReadAsStringAsync())!["app-url"]!;
        var (docId, note) = codimd.Notes.Single();
        await Assert.That(note).IsEqualTo("# Hello\n");
        await Assert.That(appUrl).StartsWith($"http://codimd/{docId}?both&wopiSrc=http%3A%2F%2Flocalhost%2Fwopi%2Ffiles%2Flocal%21");
        var query = QueryHelpers.ParseQuery(new Uri(appUrl).Query);
        var wopiSrc = query["wopiSrc"].ToString();
        var token = query["accessToken"].ToString();
        await Assert.That(query["displayName"].ToString()).IsEqualTo("Alice");

        // the app validates the token against the docid
        var bridgeUrl = $"/wopi/bridge/{docId}?WOPISrc={Uri.EscapeDataString(wopiSrc)}&access_token={token}";
        using (var head = await host.Client.SendAsync(new(HttpMethod.Head, bridgeUrl)))
        {
            await Assert.That(head.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var wrongDoc = await host.Client.SendAsync(new(HttpMethod.Head, bridgeUrl.Replace(docId, "other"))))
        {
            await Assert.That(wrongDoc.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
        }

        // the file is WOPI-locked by the bridge
        using (var list = new HttpRequestMessage(HttpMethod.Get, $"/wopi/bridge/list?apikey={TestServerHost.WopiSecret}"))
        using (var listResponse = await host.Client.SendAsync(list))
        {
            var openFiles = JsonNode.Parse(await listResponse.Content.ReadAsStringAsync())!;
            await Assert.That((string?) openFiles[wopiSrc]!["app"]).IsEqualTo("CodiMD");
        }

        // the user edits and closes the note: the save is enqueued, then performed by the save loop
        codimd.Notes[docId] = "# Hello edited\n";
        using var save = new HttpRequestMessage(HttpMethod.Post, $"{bridgeUrl}&close=true");
        save.Headers.Add("X-Efss-Bridged-App", "CodiMD");
        using (var saveResponse = await host.Client.SendAsync(save))
        {
            await Assert.That(saveResponse.StatusCode).IsEqualTo(HttpStatusCode.Accepted);
        }

        await host.Service<Bridge>().RunRound();
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("/notes.md"))).IsEqualTo("# Hello edited\n");

        // the next save request gets the outcome of the previous one
        using var again = new HttpRequestMessage(HttpMethod.Post, $"{bridgeUrl}&close=true");
        again.Headers.Add("X-Efss-Bridged-App", "CodiMD");
        using var againResponse = await host.Client.SendAsync(again);
        await Assert.That(againResponse.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await againResponse.Content.ReadAsStringAsync()).IsEqualTo("""{"message": "File saved successfully", "delay": "1.6"}""");
    }

    [Test]
    public async Task Save_UnknownApp()
    {
        await using var host = await TestServerHost.Start();
        using var save = new HttpRequestMessage(HttpMethod.Post, "/wopi/bridge/doc?WOPISrc=x&access_token=y");
        save.Headers.Add("X-Efss-Bridged-App", "Other");
        using var response = await host.Client.SendAsync(save);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        await Assert.That(await response.Content.ReadAsStringAsync()).StartsWith("""{"message": "Unknown application, could not save.""");
    }
}
