using System.Xml.Linq;

/// <summary>
/// New files from desktop Office, as on SharePoint Server SE (docs/fsshttp-notes.md): an empty file, as an application
/// creates for a new document, opens and takes the first save; and a Save As to a new name in the folder of the file,
/// which the site of the token then reaches, as it reaches no other file of the folder.
/// </summary>
public class NewFileTests
{
    [Test]
    [Arguments("new.docx")]
    [Arguments("new.xlsx")]
    [Arguments("new.pptx")]
    public async Task EmptyFile(string name)
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf(name), []);
        var (_, token) = await host.Open(name);
        var word = new WordClient(host, token);
        await Assert.That((await word.Coauth("JoinCoauthoring", new XAttribute("Timeout", 3600))).ErrorCode).IsEqualTo("Success");
        await word.Sync();
        await Assert.That(word.File).IsEmpty();

        var saved = Zips.Create(("word/document.xml", Zips.Text("first save")));
        var response = await word.Save(saved);
        await Assert.That(response).IsTypeOf<PutChangesResponse>();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf(name))).SameAs(saved)).IsTrue();
    }

    static async Task<(TestServerHost Host, string Token, WordClient Word)> Start(Mode mode = Mode.Edit)
    {
        var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf("base.docx"), Zips.Create(("word/document.xml", Zips.Text("base"))));
        await File.WriteAllBytesAsync(host.PathOf("other.docx"), Zips.Create(("word/document.xml", Zips.Text("other"))));
        var (_, token) = await host.Open("base.docx", mode);
        var word = new WordClient(host, token);
        await word.Sync();
        return (host, token, word);
    }

    static string Url(string token, string name) => $"https://localhost/fsshttp/{token}/{Uri.EscapeDataString(name)}";

    static async Task<HttpStatusCode> Head(TestServerHost host, string token, string name)
    {
        using var request = new HttpRequestMessage(HttpMethod.Head, $"/fsshttp/{token}/{Uri.EscapeDataString(name)}");
        using var response = await host.Client.SendAsync(request);
        return response.StatusCode;
    }

    /// <summary>
    /// Word checks that the name is free, then saves as it: the file is created with the content, under the exclusive
    /// lock of the Cell, which ConvertToSchemaJoinCoauth turns into a session; the site then reaches the new file by its
    /// name, for its content, its saves and its coauthoring, and the file of the token stays as it was.
    /// </summary>
    [Test]
    public async Task SaveAs()
    {
        var (host, token, _) = await Start();
        await using var _ = host;
        var original = await File.ReadAllBytesAsync(host.PathOf("base.docx"));
        await Assert.That(await Head(host, token, "Saved As.docx")).IsEqualTo(HttpStatusCode.NotFound);

        var saveAs = new WordClient(host, token);
        var content = Zips.Create(("word/document.xml", Zips.Text("saved as")));
        var results = await saveAs.SaveAs(Url(token, "Saved As.docx"), content);
        await Assert.That(results.Select(_ => _.ErrorCode)).IsEquivalentTo(["Success", "Success", "Success"]);
        await Assert.That(results[0].FssHttpB.SubResponses.Single()).IsTypeOf<PutChangesResponse>();
        await Assert.That(results[1].Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("Saved As.docx"))).SameAs(content)).IsTrue();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("base.docx"))).SameAs(original)).IsTrue();

        // the content of the new name, and a save to it under the session
        await Assert.That(await Head(host, token, "Saved As.docx")).IsEqualTo(HttpStatusCode.OK);
        var bytes = await host.Client.GetByteArrayAsync($"/fsshttp/{token}/{Uri.EscapeDataString("Saved As.docx")}");
        await Assert.That(bytes.SameAs(content)).IsTrue();
        var edited = Zips.Create(("word/document.xml", Zips.Text("edited after save as")));
        await Assert.That(await saveAs.Save(edited)).IsTypeOf<PutChangesResponse>();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("Saved As.docx"))).SameAs(edited)).IsTrue();

        // a second client of the site joins the new file, and IsOnlyClient of its ResourceID says so
        var second = new WordClient(host, token)
        {
            DocumentUrl = Url(token, "Saved As.docx")
        };
        await Assert.That((await second.Coauth("JoinCoauthoring", new XAttribute("Timeout", 3600))).ErrorCode).IsEqualTo("Success");
        var resourceId = FssHttpSite.ResourceIdOf(host.Storage.GetFileId("Saved As.docx"));
        await Assert.That(await saveAs.IsOnlyClient(resourceId.ToString())).IsFalse();
        await Assert.That(await saveAs.IsOnlyClient(FssHttpSite.ResourceIdOf(host.Storage.GetFileId("base.docx")).ToString())).IsTrue();
    }

    /// <summary>
    /// Saving as needs the right to create files next to the file, as the Save As of the WOPI clients does, and one of
    /// the formats desktop Office edits through OfficeHost.
    /// </summary>
    [Test]
    public async Task SaveAs_NotAllowed()
    {
        await using var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf("base.docx"), Zips.Create(("word/document.xml", Zips.Text("base"))));
        var fileId = host.Storage.GetFileId("base.docx");
        var token = await host.Token(TestServerHost.Access(fileId) with { CanWriteRelative = false });
        var content = Zips.Create(("word/document.xml", Zips.Text("saved as")));
        var results = await new WordClient(host, token).SaveAs(Url(token, "new.docx"), content);
        await Assert.That(results[0].ErrorCode).IsEqualTo("FileNotExistsOrCannotBeCreated");
        await Assert.That(File.Exists(host.PathOf("new.docx"))).IsFalse();

        var (_, editToken) = await host.Open("base.docx");
        results = await new WordClient(host, editToken).SaveAs(Url(editToken, "new.doc"), content);
        await Assert.That(results[0].ErrorCode).IsEqualTo("FileNotExistsOrCannotBeCreated");
        await Assert.That(File.Exists(host.PathOf("new.doc"))).IsFalse();
    }

    /// <summary>
    /// The site reaches no other file of the folder: a HEAD shows that the name is taken, so that Word asks before it
    /// replaces it, but neither its content nor its cells are served. A Save As takes over the file when it is not
    /// locked, as PutRelativeFile with an overwrite does.
    /// </summary>
    [Test]
    public async Task SaveAs_OtherFile()
    {
        var (host, token, _) = await Start();
        await using var _ = host;
        await Assert.That(await Head(host, token, "other.docx")).IsEqualTo(HttpStatusCode.OK);
        using var get = await host.Client.GetAsync($"/fsshttp/{token}/other.docx");
        await Assert.That(get.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        var reader = new WordClient(host, token)
        {
            DocumentUrl = Url(token, "other.docx")
        };
        var query = await reader.Send(WordClient.Cell("1", WordClient.Query()));
        await Assert.That(query.Single().ErrorCode).IsEqualTo("FileNotExistsOrCannotBeCreated");

        // locked by a WOPI client: not taken over
        var otherId = host.Storage.GetFileId("other.docx");
        await host.Service<IOfficeLockStore>().CompareExchange(otherId, null, new("wopi lock", DateTimeOffset.UtcNow.AddMinutes(30)));
        var content = Zips.Create(("word/document.xml", Zips.Text("saved over")));
        var results = await new WordClient(host, token).SaveAs(Url(token, "other.docx"), content);
        await Assert.That(results[0].ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");

        await host.Service<IOfficeLockStore>().CompareExchange(otherId, "wopi lock", null);
        results = await new WordClient(host, token).SaveAs(Url(token, "other.docx"), content);
        await Assert.That(results[0].ErrorCode).IsEqualTo("Success");
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("other.docx"))).SameAs(content)).IsTrue();
        await Assert.That(host.Storage.GetFileId("other.docx")).IsEqualTo(otherId);
    }

    /// <summary>
    /// A file a Save As created is deleted again when its Put Changes fails.
    /// </summary>
    [Test]
    public async Task SaveAs_FailedPut()
    {
        var (host, token, _) = await Start();
        await using var _ = host;
        var client = new WordClient(host, token)
        {
            DocumentUrl = Url(token, "new.docx")
        };
        var ids = new IdAllocator(Guid.NewGuid());
        // a Storage Index the package does not have
        var put = new PutChangesRequest(1, ids.NewId(), ExtendedGuid.Null, 0x48, new byte[3], new byte[] { 4, 0, 0 });
        var request = new FssHttpBRequest(new(Guid.NewGuid(), "msword", "win", 1), [put], new([])).ToArray();
        var result = (await client.Send(WordClient.Cell("1", request, new XAttribute("Coalesce", "true")))).Single();
        await Assert.That(result.ErrorCode).IsEqualTo("CellRequestFail");
        await Assert.That(File.Exists(host.PathOf("new.docx"))).IsFalse();
    }

    /// <summary>
    /// A Save As that takes over another file writes it, but does not read it: a query of its content in the same
    /// request fails, and when the Put Changes fails, the site does not reach the file after.
    /// </summary>
    [Test]
    public async Task SaveAs_OtherFile_NotRead()
    {
        var (host, token, _) = await Start();
        await using var _ = host;
        var original = await File.ReadAllBytesAsync(host.PathOf("other.docx"));
        var client = new WordClient(host, token)
        {
            DocumentUrl = Url(token, "other.docx")
        };
        var ids = new IdAllocator(Guid.NewGuid());
        var query = new QueryChangesRequest(1, 0, 2, new(3, CellId.Null), 3670016, Knowledge: Knowledge.Empty);
        // a Storage Index the package does not have
        var put = new PutChangesRequest(2, ids.NewId(), ExtendedGuid.Null, 0x48, new byte[3], new byte[] { 4, 0, 0 });
        var request = new FssHttpBRequest(new(Guid.NewGuid(), "msword", "win", 1), [query, put], new([])).ToArray();
        var result = (await client.Send(WordClient.Cell("1", request, new XAttribute("Coalesce", "true")))).Single();
        await Assert.That(result.ErrorCode).IsEqualTo("CellRequestFail");
        var queried = (FailedSubResponse) result.FssHttpB.SubResponses[0];
        await Assert.That(queried.Error).IsEqualTo(ResponseError.Cell(CellErrorCode.StorageWriteOnly));
        await Assert.That(result.FssHttpB.DataElements!.Elements).IsEmpty();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("other.docx"))).SameAs(original)).IsTrue();

        using var get = await host.Client.GetAsync($"/fsshttp/{token}/other.docx");
        await Assert.That(get.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// After a rename the file answers to its new name only, as on SharePoint Server SE.
    /// </summary>
    [Test]
    public async Task Rename_OldName()
    {
        var (host, token, word) = await Start();
        await using var _ = host;
        var rename = WordClient.SubRequest("FileOperation", "1", WordClient.Data(new XAttribute("FileOperation", "Rename"), new XAttribute("NewFileName", "renamed.docx")));
        await Assert.That((await word.Send(rename)).Single().ErrorCode).IsEqualTo("Success");

        word.DocumentUrl = Url(token, "base.docx");
        await Assert.That((await word.Send(WordClient.Cell("1", WordClient.Query()))).Single().ErrorCode).IsEqualTo("FileNotExistsOrCannotBeCreated");
        await Assert.That(await Head(host, token, "base.docx")).IsEqualTo(HttpStatusCode.NotFound);
        word.DocumentUrl = Url(token, "renamed.docx");
        await Assert.That((await word.Send(WordClient.Cell("1", WordClient.Query()))).Single().ErrorCode).IsEqualTo("Success");
        await Assert.That(await Head(host, token, "renamed.docx")).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// The partitions OfficeHost keeps its own state in, the coauthoring session and the files made by Save As, are not
    /// cells a client can query or put.
    /// </summary>
    [Test]
    [Arguments("B5D0E8A3-3C51-4F6B-9E2A-6C1D7F4A2E90")]
    [Arguments("0E6B2F7C-58A1-4D3E-9B64-3F8C1A2D7E15")]
    public async Task ReservedPartitions(string partition)
    {
        var (host, _, word) = await Start();
        await using var _ = host;
        var result = (await word.Send(WordClient.Cell("1", WordClient.Query(), new XAttribute("PartitionID", partition)))).Single();
        await Assert.That(result.ErrorCode).IsEqualTo("CellRequestFail");
        var failed = (FailedSubResponse) result.FssHttpB.SubResponses.Single();
        await Assert.That(failed.Error.Code).IsEqualTo((uint) CellErrorCode.InvalidPartition);
    }
}
