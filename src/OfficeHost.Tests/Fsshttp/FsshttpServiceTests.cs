using System.Xml.Linq;
using System.IO.Compression;

/// <summary>
/// The desktop site of <see cref="OfficeHostExtensions.MapFsshttp"/>: discovery, the cell storage service as Word uses it
/// (docs/fsshttp-spike.md), locks shared with WOPI, coauthoring, and the editors table.
/// </summary>
public class FsshttpServiceTests
{
    static byte[] Zip(string text) => Zips.Create(("word/document.xml", Zips.Text(text)), ("large.bin", Zips.Random(10_000, 1)));

    static async Task<(TestServerHost Host, string FileId, string Token)> Start(byte[]? content = null, Mode mode = Mode.Edit, TimeProvider? time = null)
    {
        var host = await TestServerHost.Start(time: time);
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), content ?? Zip("hello"));
        var (fileId, token) = await host.Open("doc.docx", mode);
        return (host, fileId, token);
    }

    static async Task<byte[]> FileOf(TestServerHost host) => await File.ReadAllBytesAsync(host.PathOf("doc.docx"));

    [Test]
    public async Task Options()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        using var request = new HttpRequestMessage(HttpMethod.Options, $"/fsshttp/{token}/doc.docx");
        request.Headers.Add("X-MSGETWEBURL", "t");
        using var response = await host.Client.SendAsync(request);
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(response.Headers.GetValues("X-MSFSSHTTP").Single()).IsEqualTo("1.3");
        await Assert.That(response.Headers.GetValues("X-MSWEBURL").Single()).IsEqualTo($"/fsshttp/{token}");
        await Assert.That(response.Headers.GetValues("MicrosoftSharePointTeamServices").Single()).IsEqualTo("16.0.0.0");

        using var plain = await host.Client.SendAsync(new(HttpMethod.Options, $"/fsshttp/{token}"));
        await Assert.That(plain.Headers.Contains("X-MSWEBURL")).IsFalse();
        using var invalid = await host.Client.SendAsync(new(HttpMethod.Options, "/fsshttp/invalid/doc.docx"));
        await Assert.That(invalid.StatusCode).IsEqualTo(HttpStatusCode.Unauthorized);
    }

    [Test]
    public async Task Document()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var bytes = await host.Client.GetByteArrayAsync($"/fsshttp/{token}/doc.docx");
        await Assert.That(bytes.SameAs(await FileOf(host))).IsTrue();
        using var head = await host.Client.SendAsync(new(HttpMethod.Head, $"/fsshttp/{token}/doc.docx"));
        await Assert.That(head.Content.Headers.ContentLength).IsEqualTo(bytes.Length);
        using var vti = await host.Client.GetAsync($"/fsshttp/{token}/_vti_inf.html");
        await Assert.That(vti.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        using var post = await host.Client.PostAsync($"/fsshttp/{token}/_vti_bin/webs.asmx", new StringContent(""));
        await Assert.That(post.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
        File.Delete(host.PathOf("doc.docx"));
        using var missing = await host.Client.GetAsync($"/fsshttp/{token}/doc.docx");
        await Assert.That(missing.StatusCode).IsEqualTo(HttpStatusCode.NotFound);
    }

    /// <summary>
    /// The open of Word (docs/fsshttp-spike.md): the join, the schema lock only if the join is not supported, the three
    /// Cell queries (editors table, metadata, content), the editors metadata, WhoAmI, GetDocMetaInfo, GetVersions,
    /// ServerTime and Query Access.
    /// </summary>
    [Test]
    public async Task Open()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        var results = await word.Send(
            word.Join(),
            WordClient.SubRequest("SchemaLock", "5", WordClient.Data(new XAttribute("SchemaLockRequestType", "GetLock"), new XAttribute("SchemaLockID", WordClient.SchemaLockId), new XAttribute("ClientID", word.ClientId)), "1", "OnNotSupported"),
            WordClient.Cell("8", WordClient.Query(), new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")),
            WordClient.Cell("7", WordClient.Query(), new XAttribute("PartitionID", "383ADC0B-E66E-4438-95E6-E39EF9720122")),
            WordClient.Cell("6", WordClient.Query(), new XAttribute("GetFileProps", "true")),
            WordClient.SubRequest("EditorsTable", "4", WordClient.Data(new XAttribute("EditorsTableRequestType", "UpdateEditorMetadata"), new XAttribute("ClientID", word.ClientId), new XAttribute("Key", "Platform"), new XAttribute("Value", "MQ==")), "1", "OnSuccess"),
            WordClient.SubRequest("WhoAmI", "9"),
            WordClient.SubRequest("GetDocMetaInfo", "10", null, "9"),
            WordClient.SubRequest("GetVersions", "11"),
            WordClient.SubRequest("ServerTime", "12"),
            WordClient.Cell("13", WordClient.Request(new QueryAccessRequest(1))));
        var byToken = results.ToDictionary(_ => _.Token);
        await Assert.That(byToken["1"].ErrorCode).IsEqualTo("Success");
        await Assert.That(byToken["1"].Attribute("LockType")).IsEqualTo("SchemaLock");
        await Assert.That(byToken["1"].Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(byToken["1"].Attribute("TransitionID")).IsNotNull();
        await Assert.That(byToken["5"].ErrorCode).IsEqualTo("DependentOnlyOnNotSupportedRequestGetSupported");
        await Assert.That(byToken["11"].ErrorCode).IsEqualTo("RequestNotSupported");
        foreach (var name in new[] { "4", "6", "7", "8", "9", "10", "12", "13" })
        {
            await Assert.That(byToken[name].ErrorCode).IsEqualTo("Success");
        }

        // the content, chunked as a file stream
        var content = byToken["6"].Fsshttpb;
        var query = (QueryChangesResponse) content.SubResponses.Single();
        var storage = CellStorage.Create(content.DataElements!.Elements, query.StorageIndex);
        await Assert.That(FileCell.Read(storage).Root.ToArray().SameAs(await FileOf(host))).IsTrue();
        await Assert.That(byToken["6"].Attribute("Etag")).StartsWith("\"{");
        await Assert.That(byToken["6"].Attribute("ModifiedBy")).IsEqualTo("Alice");

        // the metadata, with nothing in it: a Storage Index with a null Storage Manifest mapping
        var metadata = byToken["7"].Fsshttpb;
        var index = (StorageIndex) metadata.DataElements!.Elements.Single();
        await Assert.That(index.Entries.Single().MappingId.IsNull).IsTrue();

        // the editors table: the XML of the editors, compressed after its header
        var editors = EditorsTable(byToken["8"].Fsshttpb);
        var editor = editors.Element("Editor")!;
        await Assert.That((string?) editor.Element("CacheID")).IsEqualTo(word.ClientId);
        await Assert.That((string?) editor.Element("FriendlyName")).IsEqualTo("Alice");
        // the metadata came after the query, as in Word's open: the next query has it
        var next = EditorsTable((await word.Send(WordClient.Cell("1", WordClient.Query(), new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")))).Single().Fsshttpb);
        await Assert.That(next.Element("Editor")!.Element("Metadata")!.Element("Platform")!.Value).IsEqualTo("MQ==");

        await Assert.That(byToken["9"].Attribute("UserName")).IsEqualTo("Alice");
        await Assert.That(byToken["9"].Attribute("UserLogin")).IsEqualTo("alice");
        await Assert.That(byToken["10"].Data.Descendants().Any(_ => (string?) _.Attribute("Key") == "vti_sourcecontrollockid")).IsTrue();
        await Assert.That(long.Parse(byToken["12"].Attribute("ServerTime")!)).IsGreaterThan(0);
        var access = (QueryAccessResponse) byToken["13"].Fsshttpb.SubResponses.Single();
        await Assert.That(access.WriteAccess).IsEqualTo(ResponseError.Success);
    }

    static XElement EditorsTable(FsshttpbResponse response)
    {
        var query = (QueryChangesResponse) response.SubResponses.Single();
        var stream = FileCell.Read(CellStorage.Create(response.DataElements!.Elements, query.StorageIndex)).Root.ToArray();
        if (!stream.AsSpan(0, 8).SequenceEqual(new byte[] { 0x1A, 0x5A, 0x3A, 0x30, 0, 0, 0, 0 }))
        {
            throw new("Not the Editors Table Zip Stream Header");
        }

        using var deflate = new DeflateStream(new MemoryStream(stream, 8, stream.Length - 8), CompressionMode.Decompress);
        return XElement.Load(deflate);
    }

    /// <summary>
    /// A save of Word: a Put Changes with the expected Storage Index under an id of Word's, which writes the file; a
    /// save on a stale state fails coherency; the next query of a client brings only what changed.
    /// </summary>
    [Test]
    public async Task Save()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        await word.Send(word.Join());
        var first = await word.Sync();
        await Assert.That(word.File.SameAs(await FileOf(host))).IsTrue();

        var saved = await word.Save(Zip("hello, world"));
        await Assert.That(saved).IsTypeOf<PutChangesResponse>();
        await Assert.That(((PutChangesResponse) saved).DataElementsAdded!.Count).IsGreaterThan(0);
        await Assert.That((await FileOf(host)).SameAs(Zip("hello, world"))).IsTrue();

        // nothing new for the client that saved
        var again = await word.Sync();
        await Assert.That(again.StorageIndex).IsNotEqualTo(first.StorageIndex);

        // another client that knew the previous state
        var other = new WordClient(host, token);
        await other.Sync();
        var third = await word.Save(Zip("third"));
        await Assert.That(third).IsTypeOf<PutChangesResponse>();
        var stale = await other.Save(Zip("stale"));
        await Assert.That(((FailedSubResponse) stale).Error).IsEqualTo(ResponseError.Cell(CellErrorCode.CoherencyFailure));
        await Assert.That((await FileOf(host)).SameAs(Zip("third"))).IsTrue();
    }

    /// <summary>
    /// Two clients coauthoring (docs/fsshttp-spike.md, sixth capture): the second join starts a transition, which makes
    /// IsOnlyClient false for the first, until it marks it complete; both poll the status; a save of one comes to the
    /// other with its next query; a stale save fails coherency, and goes through after a query, as Word merges.
    /// </summary>
    [Test]
    public async Task Coauthoring()
    {
        var (host, fileId, token) = await Start();
        await using var _ = host;
        var bobToken = await host.Token(TestServerHost.Access(fileId, userId: "bob"));
        var alice = new WordClient(host, token);
        var bob = new WordClient(host, bobToken);

        var join = (await alice.Send(alice.Join())).Single();
        await Assert.That(join.Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(await IsOnlyClient(alice)).IsTrue();
        await alice.Sync();
        await bob.Sync();

        var bobJoin = (await bob.Send(bob.Join())).Single();
        await Assert.That(bobJoin.Attribute("CoauthStatus")).IsEqualTo("Coauthoring");
        await Assert.That(bobJoin.Attribute("TransitionID")).IsEqualTo(join.Attribute("TransitionID"));
        await Assert.That(await IsOnlyClient(alice)).IsFalse();
        await Assert.That((await alice.Coauth("MarkTransitionComplete")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await alice.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Coauthoring");
        await Assert.That((await bob.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Coauthoring");

        // the editors table has both
        var editors = EditorsTable((await alice.Send(WordClient.Cell("1", WordClient.Query(), new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")))).Single().Fsshttpb);
        await Assert.That(editors.Elements("Editor").Select(_ => (string) _.Element("FriendlyName")!).Order().ToList()).IsEquivalentTo(["Alice", "Bob"]);

        // Alice saves; Bob's save on the old state fails; Bob queries, merges and saves
        await Assert.That(await alice.Save(Zip("alice"))).IsTypeOf<PutChangesResponse>();
        await Assert.That(await bob.Save(Zip("bob"))).IsTypeOf<FailedSubResponse>();
        var before = bob.Content!.Elements.Count;
        await bob.Sync();
        await Assert.That(bob.File.SameAs(Zip("alice"))).IsTrue();
        await Assert.That(bob.Content.Elements.Count).IsGreaterThan(before);
        await Assert.That(await bob.Save(Zip("alice and bob"))).IsTypeOf<PutChangesResponse>();
        await alice.Sync();
        await Assert.That(alice.File.SameAs(Zip("alice and bob"))).IsTrue();
        await Assert.That((await FileOf(host)).SameAs(Zip("alice and bob"))).IsTrue();

        // Bob leaves: Alice is alone again
        await Assert.That((await bob.Coauth("ExitCoauthoring")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await alice.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(await IsOnlyClient(alice)).IsTrue();
        await Assert.That((await alice.Coauth("ExitCoauthoring")).ErrorCode).IsEqualTo("Success");
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
        // the session is gone: a second exit of a client finds no lock
        await Assert.That((await alice.Coauth("ExitCoauthoring")).ErrorCode).IsEqualTo("FileNotLockedOnServer");
    }

    static async Task<bool> IsOnlyClient(WordClient client)
    {
        var body = """<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body><IsOnlyClient xmlns="http://schemas.microsoft.com/sharepoint/soap/"><id>x</id></IsOnlyClient></s:Body></s:Envelope>""";
        using var response = await client.Post("_vti_bin/sharedaccess.asmx", new StringContent(body, Encoding.UTF8, "text/xml"));
        var xml = XDocument.Parse(await response.Content.ReadAsStringAsync());
        return bool.Parse(xml.Descendants().Single(_ => _.Name.LocalName == "IsOnlyClient" + "Result").Value);
    }

    /// <summary>
    /// The lock of desktop editing is a lock of the <see cref="IOfficeLockStore"/>: a WOPI client cannot lock the file
    /// while Word has it, and Word cannot join while a WOPI client has it. A save needs the lock it holds.
    /// </summary>
    [Test]
    public async Task Locks_Wopi()
    {
        var (host, fileId, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        await Assert.That((await word.Send(word.Join())).Single().ErrorCode).IsEqualTo("Success");
        using (var wopi = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(wopi.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        await word.Sync();
        var denied = await word.Save(Zip("other lock"), "{OTHER}");
        await Assert.That(((FailedSubResponse) denied).Error.SupplementalInfo).IsEqualTo("FileAlreadyLockedOnServer");
        await word.Coauth("ExitCoauthoring");

        using (var wopi = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(wopi.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((await word.Send(word.Join())).Single().ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await word.Coauth("CheckLockAvailability")).ErrorCode).IsEqualTo("Success");
        var blocked = await word.Save(Zip("blocked"));
        await Assert.That(((FailedSubResponse) blocked).Error.SupplementalInfo).IsEqualTo("FileAlreadyLockedOnServer");
    }

    [Test]
    public async Task Locks_Schema()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        var other = new WordClient(host, token);

        async Task<SubResult> Schema(WordClient client, string type, string schemaLockId = WordClient.SchemaLockId, params XAttribute[] attributes) =>
            (await client.Send(WordClient.SubRequest("SchemaLock", "1", WordClient.Data(new XAttribute("SchemaLockRequestType", type), new XAttribute("SchemaLockID", schemaLockId), new XAttribute("ClientID", client.ClientId), new XAttribute("Timeout", 3600), attributes)))).Single();

        var get = await Schema(word, "GetLock");
        await Assert.That(get.Attribute("LockType")).IsEqualTo("SchemaLock");
        await Assert.That((await Schema(word, "RefreshLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Schema(other, "GetLock", "{ANOTHER-SCHEMA}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Schema(other, "CheckLockAvailability", "{ANOTHER-SCHEMA}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Schema(other, "ReleaseLock")).ErrorCode).IsEqualTo("InvalidCoauthSession");
        await Assert.That((await Schema(other, "GetLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Schema(word, "ConvertToExclusive", attributes: new XAttribute("ExclusiveLockID", "{E}"))).ErrorCode).IsEqualTo("MultipleClientsInCoauthSession");
        await Assert.That((await Schema(word, "ConvertToExclusive", attributes: [new("ExclusiveLockID", "{E}"), new("ReleaseLockOnConversionToExclusiveFailure", "true")])).ErrorCode).IsEqualTo("ExitCoauthSessionAsConvertToExclusiveFailed");
        await Assert.That((await Schema(other, "ConvertToExclusive", attributes: new XAttribute("ExclusiveLockID", "{E}"))).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Schema(word, "GetLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Schema(word, "ReleaseLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Schema(word, "Unknown")).ErrorCode).IsEqualTo("InvalidArgument");
    }

    [Test]
    public async Task Locks_Exclusive()
    {
        var (host, fileId, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);

        async Task<SubResult> Exclusive(string type, string lockId = "{E}") =>
            (await word.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", type), new XAttribute("ExclusiveLockID", lockId), new XAttribute("ClientID", word.ClientId), new XAttribute("SchemaLockID", WordClient.SchemaLockId), new XAttribute("Timeout", 600))))).Single();

        await Assert.That((await Exclusive("ReleaseLock")).ErrorCode).IsEqualTo("FileNotLockedOnServer");
        await Assert.That((await Exclusive("GetLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await host.Service<IOfficeLockStore>().Get(fileId))!.Id).IsEqualTo("fsshttp-exclusive:{E}");
        await Assert.That((await Exclusive("RefreshLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Exclusive("GetLock", "{F}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive("CheckLockAvailability", "{F}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive("CheckLockAvailability")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await word.Send(word.Join())).Single().ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await word.Sync();
        await Assert.That(await word.Save(Zip("exclusive"), "{E}")).IsTypeOf<PutChangesResponse>();

        var convert = await Exclusive("ConvertToSchemaJoinCoauth");
        await Assert.That(convert.ErrorCode).IsEqualTo("Success");
        await Assert.That(convert.Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That((await host.Service<IOfficeLockStore>().Get(fileId))!.Id).IsEqualTo($"fsshttp-schema:{WordClient.SchemaLockId}");
        await Assert.That((await Exclusive("GetLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive("ConvertToSchema")).ErrorCode).IsEqualTo("FileNotLockedOnServer");
        await Assert.That((await Exclusive("ReleaseLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await word.Coauth("ConvertToExclusive", new XAttribute("ExclusiveLockID", "{E}"), new XAttribute("Timeout", 600))).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Exclusive("ConvertToSchema", "{F}")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive("ReleaseLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
    }

    /// <summary>
    /// A user who cannot write gets the content, but not the lock, the editing session or a save.
    /// </summary>
    [Test]
    public async Task ReadOnly()
    {
        var (host, _, token) = await Start(mode: Mode.ReadOnly);
        await using var _ = host;
        var word = new WordClient(host, token);
        await Assert.That((await word.Send(word.Join())).Single().ErrorCode).IsEqualTo("FileUnauthorizedAccess");
        var access = (await word.Send(WordClient.Cell("1", WordClient.Request(new QueryAccessRequest(1))))).Single();
        await Assert.That(((QueryAccessResponse) access.Fsshttpb.SubResponses.Single()).WriteAccess.Type).IsEqualTo(ResponseErrorType.HResult);
        await word.Sync();
        var save = await word.Save(Zip("denied"));
        await Assert.That(((FailedSubResponse) save).Error.SupplementalInfo).IsEqualTo("FileUnauthorizedAccess");
    }

    /// <summary>
    /// The metadata partition is Word's: what it puts, it gets back. The editors table is the server's: a Put Changes
    /// to it fails.
    /// </summary>
    [Test]
    public async Task Partitions()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        var partition = new XAttribute("PartitionID", "383ADC0B-E66E-4438-95E6-E39EF9720122");
        var empty = (await word.Send(WordClient.Cell("1", WordClient.Query(), partition))).Single().Fsshttpb;
        var emptyIndex = ((QueryChangesResponse) empty.SubResponses.Single()).StorageIndex;

        var ids = new IdAllocator(Guid.NewGuid());
        var update = FileCell.Write(Zips.Text("metadata"), null, ids, ChunkingOptions.Default);
        var put = new PutChangesRequest(1, update.Storage.StorageIndexId, emptyIndex, 0x49);
        var request = new FsshttpbRequest(null, [put], new(update.Added.Select(_ => _ with { SerialNumber = SerialNumber.Null }).ToList())).ToArray();
        var putResult = (await word.Send(WordClient.Cell("1", request, partition))).Single();
        await Assert.That(putResult.Fsshttpb.SubResponses.Single()).IsTypeOf<PutChangesResponse>();

        var back = (await word.Send(WordClient.Cell("1", WordClient.Query(), partition))).Single().Fsshttpb;
        var query = (QueryChangesResponse) back.SubResponses.Single();
        var storage = CellStorage.Create(back.DataElements!.Elements, query.StorageIndex);
        await Assert.That(Encoding.UTF8.GetString(FileCell.Read(storage).Root.ToArray())).IsEqualTo("metadata");

        // a second put on the empty state is stale
        var stale = (await word.Send(WordClient.Cell("1", request, partition))).Single();
        await Assert.That(stale.Fsshttpb.SubResponses.Single()).IsTypeOf<FailedSubResponse>();

        var editorsPut = (await word.Send(WordClient.Cell("1", request, new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")))).Single();
        await Assert.That(((FailedSubResponse) editorsPut.Fsshttpb.SubResponses.Single()).Error).IsEqualTo(ResponseError.Cell(CellErrorCode.StorageReadOnly));
    }

    /// <summary>
    /// A query with a small maximum gets part of the data elements, and the rest with the knowledge it got.
    /// </summary>
    [Test]
    public async Task Query_Partial()
    {
        var (host, _, token) = await Start(Zips.Create(("a.bin", Zips.Random(50_000, 2)), ("b.bin", Zips.Random(50_000, 3))));
        await using var _ = host;
        var word = new WordClient(host, token);
        var knowledge = Knowledge.Empty;
        var elements = new List<DataElement>();
        ExtendedGuid index = default;
        for (var round = 0; round < 20; round++)
        {
            var response = (await word.Send(WordClient.Cell("1", WordClient.Query(knowledge, 20_000)))).Single().Fsshttpb;
            var query = (QueryChangesResponse) response.SubResponses.Single();
            elements.AddRange(response.DataElements!.Elements);
            knowledge = query.Knowledge;
            index = query.StorageIndex;
            if (!query.Partial)
            {
                break;
            }
        }

        await Assert.That(elements.Count).IsGreaterThan(4);
        var storage = CellStorage.Create(elements, index);
        await Assert.That(FileCell.Read(storage).Root.ToArray().SameAs(await FileOf(host))).IsTrue();
    }

    /// <summary>
    /// The editors table sub-requests, and their limits.
    /// </summary>
    [Test]
    public async Task EditorsTable()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);

        async Task<string> Editors(string type, params XAttribute[] attributes) =>
            (await word.Send(WordClient.SubRequest("EditorsTable", "1", WordClient.Data(new XAttribute("EditorsTableRequestType", type), new XAttribute("ClientID", word.ClientId), attributes)))).Single().ErrorCode;

        await Assert.That(await Editors("UpdateEditorMetadata", new XAttribute("Key", "k"), new XAttribute("Value", "AA=="))).IsEqualTo("EditorClientIdNotFound");
        await Assert.That(await Editors("JoinEditingSession", new XAttribute("AsEditor", "true"), new XAttribute("Timeout", 3600))).IsEqualTo("Success");
        await Assert.That(await Editors("RefreshEditingSession", new XAttribute("AsEditor", "true"), new XAttribute("Timeout", 3600))).IsEqualTo("Success");
        foreach (var key in new[] { "a", "b", "c", "d" })
        {
            await Assert.That(await Editors("UpdateEditorMetadata", new XAttribute("Key", key), new XAttribute("Value", "AA=="))).IsEqualTo("Success");
        }

        await Assert.That(await Editors("UpdateEditorMetadata", new XAttribute("Key", "e"), new XAttribute("Value", "AA=="))).IsEqualTo("EditorMetadataQuotaReached");
        await Assert.That(await Editors("UpdateEditorMetadata", new XAttribute("Key", new string('k', 65)), new XAttribute("Value", "AA=="))).IsEqualTo("EditorMetadataStringExceedsLengthLimit");
        await Assert.That(await Editors("UpdateEditorMetadata")).IsEqualTo("InvalidArgument");
        await Assert.That(await Editors("RemoveEditorMetadata", new XAttribute("Key", "a"))).IsEqualTo("Success");
        await Assert.That(await Editors("Unknown")).IsEqualTo("InvalidArgument");

        var table = EditorsTable((await word.Send(WordClient.Cell("1", WordClient.Query(), new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")))).Single().Fsshttpb);
        var editor = table.Element("Editor")!;
        await Assert.That(editor.Element("Metadata")!.Elements().Select(_ => _.Name.LocalName).ToList()).IsEquivalentTo(["b", "c", "d"]);
        await Assert.That((string?) editor.Element("HasEditorPermission")).IsEqualTo("true");

        await Assert.That(await Editors("LeaveEditingSession")).IsEqualTo("Success");
        var empty = EditorsTable((await word.Send(WordClient.Cell("1", WordClient.Query(), new XAttribute("PartitionID", "7808F4DD-2385-49D6-B7CE-37ACA5E43602")))).Single().Fsshttpb);
        await Assert.That(empty.Elements().Count()).IsEqualTo(0);
    }

    /// <summary>
    /// Dependencies between sub-requests ([MS-FSSHTTP] 2.2.5.3).
    /// </summary>
    [Test]
    public async Task Dependencies()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        var results = await word.Send(
            WordClient.SubRequest("GetVersions", "1"),
            WordClient.SubRequest("WhoAmI", "2"),
            WordClient.SubRequest("ServerTime", "3", null, "1", "OnSuccess"),
            WordClient.SubRequest("ServerTime", "4", null, "2", "OnFail"),
            WordClient.SubRequest("ServerTime", "5", null, "1", "OnNotSupported"),
            WordClient.SubRequest("ServerTime", "6", null, "1", "OnSuccessOrNotSupported"),
            WordClient.SubRequest("ServerTime", "7", null, "2", "Sometimes"),
            WordClient.SubRequest("ServerTime", "8", null, "2", "OnSuccess"),
            WordClient.SubRequest("ServerTime", "9", null, "missing", "OnSuccess"),
            WordClient.SubRequest("AmIAlone", "10"),
            WordClient.SubRequest("Coauth", "11", WordClient.Data(new XAttribute("CoauthRequestType", "Unknown"), new XAttribute("ClientID", "c"), new XAttribute("SchemaLockID", "s"))),
            WordClient.SubRequest("Coauth", "12", WordClient.Data()),
            WordClient.SubRequest("ExclusiveLock", "13", WordClient.Data()),
            WordClient.SubRequest("EditorsTable", "14", WordClient.Data()),
            WordClient.SubRequest("SchemaLock", "15", WordClient.Data()),
            WordClient.SubRequest("Cell", "16"),
            WordClient.Cell("17", [1, 2, 3]));
        var codes = results.Select(_ => _.ErrorCode).ToList();
        await Assert.That(codes).IsEquivalentTo([
            "RequestNotSupported",
            "Success",
            "DependentOnlyOnSuccessRequestFailed",
            "DependentOnlyOnFailRequestSucceeded",
            "Success",
            "Success",
            "InvalidRequestDependencyType",
            "Success",
            "Success",
            "Success",
            "InvalidArgument",
            "InvalidArgument",
            "InvalidArgument",
            "InvalidArgument",
            "InvalidArgument",
            "Success",
            "InvalidArgument"
        ]);
        await Assert.That(results[9].Attribute("AmIAlone")).IsEqualTo("true");
    }

    /// <summary>
    /// MTOM requests, with the FSSHTTPB as an XOP part, as Word sends its saves; and bodies that are not SOAP.
    /// </summary>
    [Test]
    public async Task Mtom()
    {
        var (host, _, token) = await Start();
        await using var _ = host;
        var word = new WordClient(host, token);
        XNamespace xop = "http://www.w3.org/2004/08/xop/include";
        var envelope = WordClient.Envelope(
            WordClient.SubRequest("Cell", "1", WordClient.Data(new XAttribute("BinaryDataSize", 1), new XElement(xop + "Include", new XAttribute("href", "cid:http%3A%2F%2Ftempuri.org%2F0")))),
            WordClient.SubRequest("Cell", "2", WordClient.Data(new XAttribute("BinaryDataSize", 1), new XElement(xop + "Include", new XAttribute("href", "cid:missing")))));
        var content = new MultipartContent("related", "boundary-of-the-test");
        content.Headers.ContentType!.Parameters.Add(new("type", "\"application/xop+xml\""));
        var xml = new StringContent(envelope.ToString(), Encoding.UTF8, "application/xop+xml");
        xml.Headers.Add("Content-ID", "<http://tempuri.org/xml/0>");
        content.Add(xml);
        var binary = new ByteArrayContent(WordClient.Query());
        binary.Headers.Add("Content-ID", "<http://tempuri.org/0>");
        content.Add(binary);
        using var response = await word.Post("_vti_bin/cellstorage.svc/CellStorageService", content);
        var results = await WordClient.Read(response);
        await Assert.That(results[0].ErrorCode).IsEqualTo("Success");
        await Assert.That(results[0].Fsshttpb.SubResponses.Single()).IsTypeOf<QueryChangesResponse>();
        await Assert.That(results[1].ErrorCode).IsEqualTo("InvalidArgument");

        using var notSoap = await word.Post("_vti_bin/cellstorage.svc/CellStorageService", new StringContent("not xml", Encoding.UTF8, "text/xml"));
        await Assert.That(notSoap.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        using var noBody = await word.Post("_vti_bin/cellstorage.svc/CellStorageService", new StringContent("<a/>", Encoding.UTF8, "text/xml"));
        await Assert.That(noBody.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        var empty = new MultipartContent("related", "b");
        using var noParts = await word.Post("_vti_bin/cellstorage.svc/CellStorageService", empty);
        await Assert.That(noParts.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        var noBoundary = new StringContent("", Encoding.UTF8);
        noBoundary.Headers.ContentType = new("multipart/related");
        using var withoutBoundary = await word.Post("_vti_bin/cellstorage.svc/CellStorageService", noBoundary);
        await Assert.That(withoutBoundary.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
    }

    /// <summary>
    /// Clients that stop polling leave the session when their timeout passes, and the lock with the last one.
    /// </summary>
    [Test]
    public async Task Expiry()
    {
        var time = new ManualTime();
        var (host, fileId, token) = await Start(time: time);
        await using var _ = host;
        var alice = new WordClient(host, token);
        var bob = new WordClient(host, token);
        await alice.Send(alice.Join());
        await bob.Send(bob.Join());
        time.Advance(TimeSpan.FromMinutes(50));
        await Assert.That((await alice.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Coauthoring");
        // Bob's hour passes, Alice's poll kept her in
        time.Advance(TimeSpan.FromMinutes(20));
        await Assert.That((await alice.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(await IsOnlyClient(alice)).IsTrue();
        time.Advance(TimeSpan.FromHours(2));
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
        await Assert.That((await bob.Coauth("RefreshCoauthoring", new XAttribute("Timeout", 3600))).Attribute("CoauthStatus")).IsEqualTo("Alone");
    }

    /// <summary>
    /// A save that stored the cell but did not get to write the file: the next query writes it.
    /// </summary>
    [Test]
    public async Task PendingFile()
    {
        var (host, fileId, token) = await Start();
        await using var _ = host;
        var cells = host.Service<FsshttpCells>();
        var content = (await cells.GetContent(fileId))!;
        var ids = content.State.Ids();
        var update = FileCell.Write(Zip("pending"), content.State.Storage, ids, ChunkingOptions.Default);
        await cells.TryUpdate(fileId, FsshttpCells.ContentPartition, content.Version, new(update.Storage, ids.Guid, ids.Next, null));
        var word = new WordClient(host, token);
        await word.Sync();
        await Assert.That(word.File.SameAs(Zip("pending"))).IsTrue();
        await Assert.That((await FileOf(host)).SameAs(Zip("pending"))).IsTrue();
    }

    [Test]
    public async Task LaunchDesktop()
    {
        var (host, fileId, _) = await Start();
        await using var _ = host;
        File.WriteAllBytes(host.PathOf("book.xlsx"), []);
        File.WriteAllBytes(host.PathOf("deck.pptx"), []);
        File.WriteAllBytes(host.PathOf("notes.txt"), []);
        var launcher = host.Service<OfficeLauncher>();
        var context = new DefaultHttpContext();
        context.Request.Scheme = "https";
        context.Request.Host = new("office.example");
        var link = await launcher.LaunchDesktop(TestServerHost.Access(fileId), context.Request);
        await Assert.That(link).StartsWith("ms-word:ofe|u|https://office.example/fsshttp/");
        await Assert.That(link).EndsWith("/doc.docx");
        var token = link.Split('/')[^2];
        await Assert.That((await host.Service<IOfficeTokens>().Validate(token))!.Access.FileId).IsEqualTo(fileId);

        var readOnly = await launcher.LaunchDesktop(TestServerHost.Access(fileId, Mode.ReadOnly), context.Request);
        await Assert.That(readOnly).StartsWith("ms-word:ofv|u|");
        await Assert.That(await launcher.LaunchDesktop(TestServerHost.Access(host.Storage.GetFileId("book.xlsx")), context.Request)).StartsWith("ms-excel:");
        await Assert.That(await launcher.LaunchDesktop(TestServerHost.Access(host.Storage.GetFileId("deck.pptx")), context.Request)).StartsWith("ms-powerpoint:");
        await Assert.That(async () => await launcher.LaunchDesktop(TestServerHost.Access(host.Storage.GetFileId("notes.txt")), context.Request)).Throws<InvalidOperationException>();
        await Assert.That(async () => await launcher.LaunchDesktop(TestServerHost.Access("missing"), context.Request)).Throws<FileNotFoundException>();
        await Assert.That(async () => await launcher.LaunchDesktop(TestServerHost.Access(fileId))).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task LaunchDesktop_BaseUrl()
    {
        await using var host = await TestServerHost.Start(_ => _.DesktopBaseUrl = "https://desktop.example/app/");
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zip("x"));
        var link = await host.Service<OfficeLauncher>().LaunchDesktop(TestServerHost.Access(host.Storage.GetFileId("doc.docx")));
        await Assert.That(link).StartsWith("ms-word:ofe|u|https://desktop.example/app/fsshttp/");
    }
}


static class ByteArrays
{
    public static bool SameAs(this byte[] bytes, byte[] other) => bytes.AsSpan().SequenceEqual(other);
}
