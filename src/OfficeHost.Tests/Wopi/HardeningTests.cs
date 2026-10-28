/// <summary>
/// What a client cannot do beyond its access, and inputs that once failed: replacing the file of the token through a
/// Save As, writing under the lock of desktop Office, names Windows keeps, FSSHTTPB that takes far more memory than its
/// size, and the options that go with them.
/// </summary>
public class HardeningTests
{
    static async Task<(TestServerHost Host, string FileId)> Start(Action<OfficeHostOptions>? configure = null)
    {
        var host = await TestServerHost.Start(configure);
        await File.WriteAllTextAsync(host.PathOf("doc.docx"), "original");
        return (host, host.Storage.GetFileId("doc.docx"));
    }

    static async Task<HttpResponseMessage> PutRelative(TestServerHost host, string fileId, string token, string content, params (string Name, string Value)[] headers)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}?access_token={token}");
        request.Headers.Add("X-WOPI-Override", "PUT_RELATIVE");
        foreach (var (name, value) in headers)
        {
            request.Headers.TryAddWithoutValidation(name, value);
        }

        request.Content = new StringContent(content);
        return await host.Client.SendAsync(request);
    }

    /// <summary>
    /// A user who can only save a copy cannot replace the file itself with a relative target of its name, nor renew
    /// the token that way.
    /// </summary>
    [Test]
    public async Task PutRelative_SourceNotReplaced()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId, Mode.ReadOnly) with { CanWriteRelative = true });
        using var response = await PutRelative(host, fileId, token, "replaced", ("X-WOPI-RelativeTarget", "doc.docx"), ("X-WOPI-OverwriteRelativeTarget", "true"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("doc.docx"))).IsEqualTo("original");
        await Assert.That(response.Headers.GetValues("X-WOPI-ValidRelativeTarget").Single()).IsEqualTo("doc (2).docx");
    }

    /// <summary>
    /// A relative target created while the content uploads is not replaced without an overwrite.
    /// </summary>
    [Test]
    public async Task PutRelative_CreatedMeanwhile()
    {
        await using var host = await TestServerHost.Start(wrapStorage: _ => new CreatingStorage(_));
        await File.WriteAllTextAsync(host.PathOf("doc.docx"), "original");
        var (fileId, token) = await host.Open("doc.docx");
        using var response = await PutRelative(host, fileId, token, "mine", ("X-WOPI-RelativeTarget", "race.docx"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("race.docx"))).IsEqualTo("theirs");
    }

    /// <summary>
    /// A storage where another client creates the relative target just before it is created.
    /// </summary>
    sealed class CreatingStorage(IOfficeStorage inner) :
        ForwardingStorage(inner)
    {
        public override async Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default)
        {
            if (name == "race.docx")
            {
                await base.CreateSibling(fileId, name, new MemoryStream("theirs"u8.ToArray()), false, cancel);
            }

            return await base.CreateSibling(fileId, name, content, overwrite, cancel);
        }
    }

    /// <summary>
    /// The lock of desktop Office, which GetLock shows, is not one a WOPI client can use: neither to write, to lock again,
    /// nor to unlock and lock with its own.
    /// </summary>
    [Test]
    public async Task DesktopLock_NotUsableByWopi()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId));
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zips.Create(("word/document.xml", Zips.Text("desktop"))));
        var word = new WordClient(host, token);
        await Assert.That((await word.Send(word.Join())).Single().ErrorCode).IsEqualTo("Success");
        var desktopLock = CoauthSessions.SchemaPrefix + WordClient.SchemaLockId;

        using (var put = await host.PutFile(fileId, token, "wopi", ("X-WOPI-Lock", desktopLock)))
        {
            await Assert.That(put.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        using (var relock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", desktopLock)))
        {
            await Assert.That(relock.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
        }

        using (var unlockAndRelock = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "mine"), ("X-WOPI-OldLock", desktopLock)))
        {
            await Assert.That(unlockAndRelock.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        using (var rename = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-RequestedName", "renamed"), ("X-WOPI-Lock", desktopLock)))
        {
            await Assert.That(rename.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        await Assert.That((await host.Service<IOfficeLockStore>().Get(fileId))!.Id).IsEqualTo(desktopLock);
    }

    /// <summary>
    /// The exclusive lock of desktop Office is its user's: another writer who learns its id cannot release it, take it
    /// over, nor save under it.
    /// </summary>
    [Test]
    public async Task ExclusiveLock_OfItsUser()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zips.Create(("word/document.xml", Zips.Text("desktop"))));
        var alice = new WordClient(host, await host.Token(TestServerHost.Access(fileId)));
        var bob = new WordClient(host, await host.Token(TestServerHost.Access(fileId, userId: "bob")));

        static async Task<SubResult> Exclusive(WordClient client, string type) =>
            (await client.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", type), new XAttribute("ExclusiveLockID", "{E}"), new XAttribute("Timeout", 600))))).Single();

        await Assert.That((await Exclusive(alice, "GetLock")).ErrorCode).IsEqualTo("Success");
        await Assert.That((await Exclusive(bob, "RefreshLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive(bob, "ReleaseLock")).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await bob.Sync();
        var denied = await bob.Save(Zips.Create(("word/document.xml", Zips.Text("bob"))), "{E}", exclusive: true);
        await Assert.That(((FailedSubResponse) denied).Error.SupplementalInfo).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await Exclusive(alice, "ReleaseLock")).ErrorCode).IsEqualTo("Success");
    }

    /// <summary>
    /// A reader cannot release an exclusive lock.
    /// </summary>
    [Test]
    public async Task ExclusiveLock_ReleaseNeedsWrite()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var reader = new WordClient(host, await host.Token(TestServerHost.Access(fileId, Mode.ReadOnly)));
        var result = (await reader.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", "ReleaseLock"), new XAttribute("ExclusiveLockID", "{E}"))))).Single();
        await Assert.That(result.ErrorCode).IsEqualTo("FileUnauthorizedAccess");
    }

    /// <summary>
    /// A Timeout out of the range of an int is an invalid argument of its sub-request, not a failure of the request.
    /// </summary>
    [Test]
    public async Task Timeout_OutOfRange()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var word = new WordClient(host, await host.Token(TestServerHost.Access(fileId)));
        var result = await word.Coauth("JoinCoauthoring", new XAttribute("Timeout", "99999999999"));
        await Assert.That(result.ErrorCode).IsEqualTo("InvalidArgument");
    }

    /// <summary>
    /// Far more sub-requests than Office sends are refused before any runs.
    /// </summary>
    [Test]
    public async Task SubRequests_TooMany()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var word = new WordClient(host, await host.Token(TestServerHost.Access(fileId)));
        var subRequests = Enumerable.Range(0, 1000).Select(_ => WordClient.SubRequest("WhoAmI", _.ToString(), WordClient.Data())).ToArray();
        await Assert.That(async () => await word.Send(subRequests)).Throws<Exception>();
    }

    /// <summary>
    /// An array of a million null Extended GUIDs, a byte each, would take 20 MB: it fails rather than allocates.
    /// </summary>
    [Test]
    public async Task FssHttpB_AllocationBudget()
    {
        var count = 1_000_000;
        var bytes = new byte[CompactUInt64.Size((ulong) count) + count];
        var writer = new ArrayBufferWriter<byte>();
        var output = new FssHttpBWriter(writer);
        output.WriteCompactUInt64((ulong) count);
        writer.WrittenSpan.CopyTo(bytes);
        await Assert.That(() => new FssHttpBReader(bytes).ReadExtendedGuidArray()).Throws<FssHttpBException>();

        // a small one is read
        await Assert.That(new FssHttpBReader(new byte[] { (3 << 1) | 1, 0, 0, 0 }).ReadExtendedGuidArray().Length).IsEqualTo(3);
    }

    [Test]
    [Arguments("CON.docx")]
    [Arguments("nul")]
    [Arguments("com1.txt")]
    [Arguments("LPT9 .docx")]
    public async Task FileNames_DeviceNames(string name)
    {
        await Assert.That(FileNames.IsLegal(name)).IsFalse();
        await Assert.That(FileNames.IsLegal(FileNames.Legal(name))).IsTrue();
    }

    [Test]
    public async Task FileNames_LoneSurrogate()
    {
        await Assert.That(FileNames.IsLegal("a\uD800b.docx")).IsFalse();
        await Assert.That(FileNames.Legal("a\uD800b.docx")).IsEqualTo("a_b.docx");
        await Assert.That(FileNames.IsLegal("a😀b.docx")).IsTrue();
        await Assert.That(FileNames.Legal("a😀b.docx")).IsEqualTo("a😀b.docx");
    }

    /// <summary>
    /// A name that is not UTF-7, which is ASCII, is kept as the client sent it.
    /// </summary>
    [Test]
    public async Task FileNames_NotUtf7()
    {
        await Assert.That(FileNames.DecodeUtf7("Résumé.docx")).IsEqualTo("Résumé.docx");
        await Assert.That(FileNames.DecodeUtf7("R+AOk-sum+AOk-.docx")).IsEqualTo("Résumé.docx");
    }

    /// <summary>
    /// A rename to a name Windows keeps for a device is an invalid name, not a server error.
    /// </summary>
    [Test]
    public async Task Rename_DeviceName()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId));
        using var response = await host.Post(fileId, token, "RENAME_FILE", ("X-WOPI-RequestedName", "CON"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.BadRequest);
    }

    /// <summary>
    /// A folder of the name a Save As suggests is skipped, as a file of that name is.
    /// </summary>
    [Test]
    public async Task PutRelative_FolderOfTheName()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        Directory.CreateDirectory(host.PathOf("copy.docx"));
        var token = await host.Token(TestServerHost.Access(fileId));
        using var response = await PutRelative(host, fileId, token, "copy", ("X-WOPI-SuggestedTarget", "copy.docx"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("copy (2).docx"))).IsEqualTo("copy");
    }

    /// <summary>
    /// Tokens signed with a previous key stay valid, those of an unknown key do not.
    /// </summary>
    [Test]
    public async Task SigningKey_Rotation()
    {
        var oldKey = new string('o', 32);
        await using var before = await TestServerHost.Start(_ => _.SigningKey = oldKey);
        var access = TestServerHost.Access("file");
        var oldToken = await before.Token(access);

        await using var after = await TestServerHost.Start(
            _ =>
            {
                _.SigningKey = new('n', 32);
                _.PreviousSigningKeys.Add(oldKey);
            });
        var tokens = after.Service<IOfficeTokens>();
        await Assert.That((await tokens.Validate(oldToken))!.Access).IsEqualTo(access);
        await using var other = await TestServerHost.Start(_ => _.SigningKey = new('x', 32));
        await Assert.That(await other.Service<IOfficeTokens>().Validate(oldToken)).IsNull();
    }

    /// <summary>
    /// Collabora Online's save of a file changed since its CheckFileInfo is a conflict it asks the user about.
    /// </summary>
    [Test]
    public async Task PutFile_CollaboraTimestamp()
    {
        var (host, fileId) = await Start();
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId));
        using (var locked = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var stale = await host.PutFile(fileId, token, "stale", ("X-WOPI-Lock", "L"), ("X-COOL-WOPI-Timestamp", "2000-01-01T00:00:00.0000000Z")))
        {
            await Assert.That(stale.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
            await Assert.That(JsonNode.Parse(await stale.Content.ReadAsStringAsync())!["COOLStatusCode"]!.GetValue<int>()).IsEqualTo(1010);
        }

        var lastModified = (await host.Storage.GetFile(fileId))!.LastModified.UtcDateTime.ToString("O", CultureInfo.InvariantCulture);
        using var saved = await host.PutFile(fileId, token, "saved", ("X-WOPI-Lock", "L"), ("X-COOL-WOPI-Timestamp", lastModified));
        await Assert.That(saved.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await File.ReadAllTextAsync(host.PathOf("doc.docx"))).IsEqualTo("saved");
    }

    /// <summary>
    /// OnFileSaved hears of the saves, with the editors the WOPI client names, and its failure does not fail the save.
    /// </summary>
    [Test]
    public async Task OnFileSaved()
    {
        var saves = new List<FileSavedContext>();
        var (host, fileId) = await Start(
            _ => _.OnFileSaved = saved =>
            {
                saves.Add(saved);
                throw new InvalidOperationException("audit down");
            });
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId));
        using (var locked = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "L")))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var saved = await host.PutFile(fileId, token, "saved", ("X-WOPI-Lock", "L"), ("X-WOPI-Editors", "alice, bob"), ("X-COOL-WOPI-IsAutosave", "true"));
        await Assert.That(saved.StatusCode).IsEqualTo(HttpStatusCode.OK);
        var save = saves.Single();
        await Assert.That(save.File.Id).IsEqualTo(fileId);
        await Assert.That(save.Editors).IsEquivalentTo(["alice", "bob"]);
        await Assert.That(save.IsAutosave).IsTrue();
        await Assert.That(save.IsExitSave).IsFalse();
    }

    [Test]
    public async Task GetShareUrl()
    {
        var (host, fileId) = await Start(_ => _.ShareUrl = (file, _, type) => Task.FromResult<string?>($"https://app/share/{file.Id}?{type}"));
        await using var _ = host;
        var token = await host.Token(TestServerHost.Access(fileId));
        using var response = await host.Post(fileId, token, "GET_SHARE_URL", ("X-WOPI-UrlType", "ReadOnly"));
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(JsonNode.Parse(await response.Content.ReadAsStringAsync())!["ShareUrl"]!.GetValue<string>()).IsEqualTo($"https://app/share/{fileId}?ReadOnly");
        using var unknown = await host.Post(fileId, token, "GET_SHARE_URL", ("X-WOPI-UrlType", "Other"));
        await Assert.That(unknown.StatusCode).IsEqualTo(HttpStatusCode.NotImplemented);
    }
}
