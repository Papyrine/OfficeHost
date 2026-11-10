/// <summary>
/// The replay of recorded exchanges to an OfficeHost (docs/replay.md).
/// </summary>
public class ReplayTests
{
    /// <summary>
    /// The folder of a recording to replay, with the report written in it.
    /// </summary>
    const string recordingVariable = "OFFICEHOST_REPLAY";

    /// <summary>
    /// A session of two Word clients recorded with OfficeHost and replayed to another: the same answers, the Cells with
    /// the ids of the first translated (<see cref="ReplayCells"/>), and the same file after each Cell of the content:
    /// the two queries, the save, and the query of the other client after it.
    /// </summary>
    [Test]
    public async Task OfficeHostAgainstItself()
    {
        using var folder = new TempDirectory();
        await using (var recorded = await TestServerHost.Start())
        {
            var original = Zips.Create(
                ("word/document.xml", Zips.Random(3000, 1)),
                ("word/styles.xml", Zips.Random(3000, 2)));
            await File.WriteAllBytesAsync(recorded.PathOf("doc.docx"), original);
            var (_, token) = await recorded.Open("doc.docx");
            recorded.Record(folder);
            var word = new WordClient(recorded, token);
            var other = new WordClient(recorded, token);

            var join = (await word.Send(word.Join())).Single();
            var transition = join.Attribute("TransitionID")!;
            await word.Sync();
            await word.IsOnlyClient(transition);
            await word.Send(WordClient.SubRequest("AmIAlone", "1", WordClient.Data(new XAttribute("TransitionID", transition))));
            await other.Send(other.Join());
            await other.Sync();
            await word.Coauth("GetCoauthoringStatus");
            await word.Send(WordClient.SubRequest("WhoAmI", "1"), WordClient.SubRequest("ServerTime", "2"), WordClient.SubRequest("LockStatus", "3"));
            await word.Save(Zips.Create(("word/document.xml", Zips.Random(3000, 3)), ("word/styles.xml", Zips.Random(3000, 2))));
            await other.Sync();
            await other.Coauth("ExitCoauthoring");
            await word.IsOnlyClient(transition);
            await word.Coauth("ExitCoauthoring");
        }

        var result = await Replayer.Run(Exchange.ReadFolder(folder));
        Console.WriteLine(result);

        await Assert.That(result.Lines[0]).IsEqualTo("Seeded doc.docx, " + result.Lines[0].Split(", ")[1]);
        await Assert.That(result.Skipped).IsEqualTo(0);
        await Assert.That(result.Different).IsEmpty();
        await Assert.That(result.Same).IsEqualTo(15);
        await Assert.That(result.Lines).Contains("#9 Cell PutChanges (translated): same Success");
        await Assert.That(result.Lines).Contains("#10 Cell QueryChanges (translated): same Success");
        await Assert.That(result.Contents).IsEqualTo(4);

        // a recorded answer that OfficeHost does not give
        var status = Path.Combine(folder, "0007.response.http");
        await File.WriteAllTextAsync(status, (await File.ReadAllTextAsync(status)).Replace("CoauthStatus=\"Coauthoring\"", "CoauthStatus=\"Alone\""));
        var tampered = await Replayer.Run(Exchange.ReadFolder(folder));
        await Assert.That(tampered.Different["Coauth GetCoauthoringStatus"]).IsEqualTo(1);
        await Assert.That(tampered.Lines).Contains("    @CoauthStatus: Alone / Coauthoring");
    }

    /// <summary>
    /// A client that knew the document before the recording, as Word opening it from Office's cache: its query has
    /// knowledge the replay never saw, and its save changes data elements the recording does not have, so neither is
    /// translated, and both say why.
    /// </summary>
    [Test]
    public async Task OfficeHostFromCache()
    {
        using var folder = new TempDirectory();
        var files = Directory.CreateDirectory(Path.Combine(folder, "files")).FullName;
        await using (var recorded = await TestServerHost.Start())
        {
            await File.WriteAllBytesAsync(recorded.PathOf("doc.docx"), Zips.Create(("word/document.xml", Zips.Random(3000, 1))));
            var (_, token) = await recorded.Open("doc.docx");
            var word = new WordClient(recorded, token);
            await word.Sync();
            await word.Save(Zips.Create(("word/document.xml", Zips.Random(3000, 2))));
            // the file as the recording starts
            File.Copy(recorded.PathOf("doc.docx"), Path.Combine(files, "doc.docx"));

            recorded.Record(folder);
            await word.Sync();
            await word.Save(Zips.Create(("word/document.xml", Zips.Random(3000, 3))));
        }

        var result = await Replayer.Run(Exchange.ReadFolder(folder), files);
        Console.WriteLine(result);

        await Assert.That(result.Lines[0]).StartsWith("Seeded doc.docx");
        await Assert.That(result.Lines).Contains(_ => _.StartsWith("#1 Cell QueryChanges (Office's cache)"));
        await Assert.That(result.Lines).Contains(_ => _.StartsWith("#2 Cell PutChanges (Office's cache)"));
        await Assert.That(result.Contents).IsEqualTo(0);
    }

    /// <summary>
    /// An MTOM body, as Office sends it, with a line break before the first boundary, written again with another
    /// envelope: the binary parts as they were.
    /// </summary>
    [Test]
    public async Task Mtom_WrittenAgain()
    {
        var binary = Enumerable.Range(0, 300).Select(_ => (byte) _).ToArray();
        var envelope = """<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><RequestCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/"><Request Url="http://sharepoint/doc.docx"><SubRequest Type="Cell" SubRequestToken="1"><SubRequestData><xop:Include href="cid:http%3A%2F%2Ftempuri.org%2F1%2F0" xmlns:xop="http://www.w3.org/2004/08/xop/include" /></SubRequestData></SubRequest></Request></RequestCollection></s:Body></s:Envelope>""";
        var boundary = "urn:uuid:0d9b1a9e-0000-4000-8000-000000000001";
        byte[] body =
        [
            .. Encoding.UTF8.GetBytes($"\r\n--{boundary}\r\nContent-ID: <http://tempuri.org/0>\r\nContent-Type: application/xop+xml;charset=utf-8;type=\"text/xml; charset=utf-8\"\r\n\r\n{envelope}\r\n--{boundary}\r\nContent-ID: <http://tempuri.org/1/0>\r\nContent-Type: application/octet-stream\r\n\r\n"),
            .. binary,
            .. Encoding.UTF8.GetBytes($"\r\n--{boundary}--\r\n")
        ];
        var contentType = $"multipart/related; type=\"application/xop+xml\"; boundary=\"{boundary}\"; start=\"<http://tempuri.org/0>\"";
        var message = new HttpMessage("POST /_vti_bin/cellstorage.svc/CellStorageService HTTP/1.1", [("Content-Type", contentType)], body);

        var read = SoapBody.Read(message);
        var data = read.Document.Descendants(SoapBody.Soap + "SubRequestData").Single();
        await Assert.That(read.Binary(data)!.SequenceEqual(binary)).IsTrue();

        var document = new XDocument(read.Document);
        document.Descendants(SoapBody.Soap + "Request").Single().SetAttributeValue("Url", "http://localhost/fsshttp/token/doc.docx");
        var again = SoapBody.Read(message with { Body = read.With(document) });
        var request = again.Document.Descendants(SoapBody.Soap + "Request").Single();
        await Assert.That((string?) request.Attribute("Url")).IsEqualTo("http://localhost/fsshttp/token/doc.docx");
        await Assert.That(again.Binary(request.Descendants(SoapBody.Soap + "SubRequestData").Single())!.SequenceEqual(binary)).IsTrue();
    }

    /// <summary>
    /// Where SharePoint Server 2016 answers otherwise than SE, which OfficeHost follows, the differences are known (as
    /// recorded of Word against a 2016 farm): its editors table without JoinCoauthTime, and without HasEditorPermission
    /// after an update of the metadata, HaveOnlyDemotionChanges with every Cell of the content, InvalidCoauthSession for
    /// MarkTransitionComplete out of the session. An editor more or less, or another error, is still different.
    /// </summary>
    [Test]
    public async Task KnownDifferences_SharePoint2016()
    {
        const string id = "{1777CAA2-7C34-4C6C-BADB-6B61E134B644}";
        var sharePoint = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id.ToLowerInvariant()}</CacheID><HasEditorPermission /><Metadata><SupportsRename>AQ==</SupportsRename></Metadata></Editor>"))!;
        var officeHost = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id}</CacheID><HasEditorPermission>true</HasEditorPermission><Metadata><JoinCoauthTime>AA==</JoinCoauthTime><SupportsRename>AQ==</SupportsRename></Metadata></Editor>"))!;
        await Assert.That(sharePoint["editors"]).IsEqualTo($"{id} metadata SupportsRename");
        await Assert.That(officeHost["editors"]).IsEqualTo($"{id} metadata SupportsRename");
        await Assert.That(officeHost["editors JoinCoauthTime"]).IsEqualTo(id);
        await Assert.That(Outcome("Cell QueryChanges (translated)", sharePoint, officeHost)).IsEqualTo(ShadowOutcome.KnownDifference);

        var another = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor><Editor><CacheID>{{00000000-0000-4000-8000-000000000002}}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor>"))!;
        await Assert.That(Outcome("Cell QueryChanges (translated)", sharePoint, another)).IsEqualTo(ShadowOutcome.Different);

        // two editors, of which one has updated its metadata: SharePoint 2016 has the permission of the other only
        const string other = "{00000000-0000-4000-8000-000000000002}";
        var mixed = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id}</CacheID><HasEditorPermission /></Editor><Editor><CacheID>{other}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor>"))!;
        var both = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor><Editor><CacheID>{other}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor>"))!;
        var refused = ShadowComparer.EditorsOf(EditorsTable($"<Editor><CacheID>{id}</CacheID><HasEditorPermission>true</HasEditorPermission></Editor><Editor><CacheID>{other}</CacheID><HasEditorPermission>false</HasEditorPermission></Editor>"))!;
        await Assert.That(mixed["editors HasEditorPermission"]).IsEqualTo($"{other}=true; {id}=");
        await Assert.That(Outcome("Cell QueryChanges (translated)", mixed, both)).IsEqualTo(ShadowOutcome.KnownDifference);
        await Assert.That(Outcome("Cell QueryChanges (translated)", mixed, refused)).IsEqualTo(ShadowOutcome.Different);
        await Assert.That(Outcome("Cell QueryChanges", new() { ["ErrorCode"] = "Success" }, new() { ["ErrorCode"] = "Success", ["editors"] = "" })).IsEqualTo(ShadowOutcome.KnownDifference);

        foreach (var label in new[] { "Cell QueryChanges", "Cell QueryAccess", "Cell PutChanges (translated)" })
        {
            await Assert.That(Outcome(label, new() { ["@HaveOnlyDemotionChanges"] = "False" }, [])).IsEqualTo(ShadowOutcome.KnownDifference);
        }

        await Assert.That(Outcome("Cell QueryAccess", new() { ["@HaveOnlyDemotionChanges"] = "True" }, [])).IsEqualTo(ShadowOutcome.Different);
        await Assert.That(Outcome("Coauth MarkTransitionComplete", new() { ["ErrorCode"] = "InvalidCoauthSession" }, new() { ["ErrorCode"] = "Success" })).IsEqualTo(ShadowOutcome.KnownDifference);
        await Assert.That(Outcome("Coauth MarkTransitionComplete", new() { ["ErrorCode"] = "Success" }, new() { ["ErrorCode"] = "InvalidCoauthSession" })).IsEqualTo(ShadowOutcome.OfficeHostFailed);

        static ShadowOutcome Outcome(string label, Dictionary<string, string> sharePoint, Dictionary<string, string> officeHost) =>
            ShadowComparer.Outcome(label, ShadowComparer.Differences(sharePoint, officeHost), ShadowRule.Defaults, out _);
    }

    /// <summary>
    /// An editors table as the servers send it in its partition: the header, then the XML deflated.
    /// </summary>
    static byte[] EditorsTable(string editors)
    {
        using var stream = new MemoryStream();
        stream.Write([0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00]);
        using (var deflate = new DeflateStream(stream, CompressionLevel.Optimal, true))
        {
            deflate.Write(Encoding.UTF8.GetBytes($"<EditorsTable>{editors}</EditorsTable>"));
        }

        return stream.ToArray();
    }

    /// <summary>
    /// Replays the recording of the folder of OFFICEHOST_REPLAY (with the documents to start from in its <c>files</c>
    /// folder, if any), and writes the report to its <c>replay.txt</c>.
    /// </summary>
    [Test]
    public async Task Recording()
    {
        var folder = Environment.GetEnvironmentVariable(recordingVariable);
        if (string.IsNullOrEmpty(folder))
        {
            Skip.Test($"No recording: set {recordingVariable} to its folder");
        }

        var result = await Replayer.Run(Exchange.ReadFolder(folder), Path.Combine(folder, "files"), folder);
        var report = Path.Combine(folder, "replay.txt");
        await File.WriteAllTextAsync(report, result.ToString());
        Console.WriteLine(result);
        Console.WriteLine($"Written to {report}");
    }
}
