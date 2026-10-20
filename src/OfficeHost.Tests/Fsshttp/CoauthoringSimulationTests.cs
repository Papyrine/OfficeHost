using System.Xml.Linq;

/// <summary>
/// Sessions of several clients over time, with clients that do what Word does (<see cref="WordClient"/>): saves in turn,
/// saves at the same time, a client that disappears, a file changed outside the session, and a file passed between
/// desktop Office and a WOPI client.
/// </summary>
public class CoauthoringSimulationTests
{
    static byte[] Zip(string text) => Zips.Create(("word/document.xml", Zips.Text(text)), ("large.bin", Zips.Random(10_000, 1)));

    static async Task<(TestServerHost Host, string FileId, WordClient Alice, WordClient Bob)> Start(TimeProvider? time = null, Action<IServiceCollection>? services = null)
    {
        var host = await TestServerHost.Start(time: time, services: services);
        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zip("start"));
        var (fileId, token) = await host.Open("doc.docx");
        var bobToken = await host.Token(TestServerHost.Access(fileId, userId: "bob"));
        return (host, fileId, new(host, token), new(host, bobToken));
    }

    static Task<byte[]> FileOf(TestServerHost host) =>
        File.ReadAllBytesAsync(host.PathOf("doc.docx"));

    static async Task<int> StoredElements(TestServerHost host, string fileId) =>
        (await host.Service<FsshttpCells>().GetContent(fileId))!.State.Storage!.Elements.Count;

    /// <summary>
    /// Two clients saving in turn, each querying first as Word does when coauthoring: every save goes through, each sees
    /// the other's, and the stored cell does not grow with the revisions.
    /// </summary>
    [Test]
    public async Task SavesInTurn()
    {
        var (host, fileId, alice, bob) = await Start();
        await using var _ = host;
        await alice.Send(alice.Join());
        await bob.Send(bob.Join());
        int? stored = null;
        for (var round = 0; round < 10; round++)
        {
            var (writer, reader) = (alice, bob);
            if (round % 2 == 1)
            {
                (writer, reader) = (bob, alice);
            }

            await writer.Sync();
            var text = $"round {round}";
            await Assert.That(await writer.Save(Zip(text))).IsTypeOf<PutChangesResponse>();
            await reader.Sync();
            await Assert.That(reader.File.SameAs(Zip(text))).IsTrue();
            await Assert.That((await FileOf(host)).SameAs(Zip(text))).IsTrue();

            var count = await StoredElements(host, fileId);
            stored ??= count;
            await Assert.That(count).IsEqualTo(stored.Value);
        }
    }

    /// <summary>
    /// Two clients saving on the same state at the same time: one wins, the other fails coherency, and the file is the
    /// winner's. The loser queries and saves again.
    /// </summary>
    [Test]
    public async Task SavesAtOnce()
    {
        var (host, _, alice, bob) = await Start();
        await using var _ = host;
        await alice.Send(alice.Join());
        await bob.Send(bob.Join());
        await alice.Sync();
        await bob.Sync();

        var results = await Task.WhenAll(alice.Save(Zip("alice")), bob.Save(Zip("bob")));
        var succeeded = results.Count(_ => _ is PutChangesResponse);
        await Assert.That(succeeded).IsEqualTo(1);
        var failed = (FailedSubResponse) results.Single(_ => _ is FailedSubResponse);
        await Assert.That(failed.Error).IsEqualTo(ResponseError.Cell(CellErrorCode.CoherencyFailure));

        var (winner, loser, text) = (alice, bob, "alice");
        if (results[1] is PutChangesResponse)
        {
            (winner, loser, text) = (bob, alice, "bob");
        }

        await Assert.That((await FileOf(host)).SameAs(Zip(text))).IsTrue();
        await loser.Sync();
        await Assert.That(loser.File.SameAs(Zip(text))).IsTrue();
        await Assert.That(await loser.Save(Zip("both"))).IsTypeOf<PutChangesResponse>();
        await winner.Sync();
        await Assert.That(winner.File.SameAs(Zip("both"))).IsTrue();
    }

    /// <summary>
    /// The file changed by the application during the session: the next save of a client that did not query fails
    /// coherency, and its query brings only what changed.
    /// </summary>
    [Test]
    public async Task ExternalChange()
    {
        var (host, fileId, alice, _) = await Start();
        await using var _ = host;
        await alice.Send(alice.Join());
        await alice.Sync();
        var before = await StoredElements(host, fileId);

        await File.WriteAllBytesAsync(host.PathOf("doc.docx"), Zip("changed outside"));
        var stale = await alice.Save(Zip("alice"));
        await Assert.That(((FailedSubResponse) stale).Error).IsEqualTo(ResponseError.Cell(CellErrorCode.CoherencyFailure));
        await Assert.That((await FileOf(host)).SameAs(Zip("changed outside"))).IsTrue();

        var known = alice.Content!.Elements.Count;
        await alice.Sync();
        await Assert.That(alice.File.SameAs(Zip("changed outside"))).IsTrue();
        // the large item is kept: fewer new elements than the whole file has
        await Assert.That(alice.Content.Elements.Count - known).IsLessThan(before);
        await Assert.That(await alice.Save(Zip("alice"))).IsTypeOf<PutChangesResponse>();
        await Assert.That((await FileOf(host)).SameAs(Zip("alice"))).IsTrue();
    }

    /// <summary>
    /// A client that disappears without exiting: the other stays in the session; once both stop, the lock expires and a
    /// WOPI client can lock the file; the client that comes back cannot join or save until it is released.
    /// </summary>
    [Test]
    public async Task ClientDisappears()
    {
        var time = new ManualTime();
        var (host, fileId, alice, bob) = await Start(time);
        await using var _ = host;
        await alice.Send(alice.Join());
        await bob.Send(bob.Join());
        await alice.Sync();
        await bob.Sync();

        // Bob is gone; Alice polls and saves for two hours
        for (var poll = 0; poll < 8; poll++)
        {
            time.Advance(TimeSpan.FromMinutes(15));
            await alice.Coauth("RefreshCoauthoring", new XAttribute("Timeout", 3600));
        }

        await Assert.That((await alice.Coauth("GetCoauthoringStatus")).Attribute("CoauthStatus")).IsEqualTo("Alone");
        await Assert.That(await alice.Save(Zip("alice"))).IsTypeOf<PutChangesResponse>();

        // Alice is gone too
        time.Advance(TimeSpan.FromHours(2));
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
        var wopiToken = await host.Token(TestServerHost.Access(fileId, userId: "carol"));
        using (var wopi = await host.Post(fileId, wopiToken, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(wopi.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        await Assert.That((await bob.Coauth("RefreshCoauthoring", new XAttribute("Timeout", 3600))).ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await bob.Send(bob.Join())).Single().ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        var save = await bob.Save(Zip("bob"));
        await Assert.That(((FailedSubResponse) save).Error.SupplementalInfo).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That((await FileOf(host)).SameAs(Zip("alice"))).IsTrue();
    }

    /// <summary>
    /// A file passed between a WOPI client and desktop Office: each sees what the other saved, and desktop Office
    /// downloads only what changed.
    /// </summary>
    [Test]
    public async Task WopiHandOff()
    {
        var (host, fileId, alice, _) = await Start();
        await using var _ = host;
        var wopiToken = await host.Token(TestServerHost.Access(fileId, userId: "carol"));

        await alice.Send(alice.Join());
        await alice.Sync();
        await Assert.That(await alice.Save(Zip("desktop 1"))).IsTypeOf<PutChangesResponse>();
        await alice.Coauth("ExitCoauthoring");

        await WopiEdit(host, fileId, wopiToken, Zip("wopi 1"));
        await Assert.That((await FileOf(host)).SameAs(Zip("wopi 1"))).IsTrue();

        await alice.Send(alice.Join());
        var known = alice.Content!.Elements.Count;
        await alice.Sync();
        await Assert.That(alice.File.SameAs(Zip("wopi 1"))).IsTrue();
        await Assert.That(alice.Content.Elements.Count - known).IsLessThan(known);
        await Assert.That(await alice.Save(Zip("desktop 2"))).IsTypeOf<PutChangesResponse>();

        // the desktop session holds the lock: the WOPI client cannot edit until it ends
        using (var locked = await host.Post(fileId, wopiToken, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.Conflict);
        }

        await alice.Coauth("ExitCoauthoring");
        await WopiEdit(host, fileId, wopiToken, Zip("wopi 2"));
        using var contents = await host.Client.GetAsync($"/wopi/files/{fileId}/contents?access_token={wopiToken}");
        await Assert.That((await contents.Content.ReadAsByteArrayAsync()).SameAs(Zip("wopi 2"))).IsTrue();
    }

    static async Task WopiEdit(TestServerHost host, string fileId, string token, byte[] content)
    {
        using (var locked = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(locked.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using (var request = new HttpRequestMessage(HttpMethod.Post, $"/wopi/files/{fileId}/contents?access_token={token}"))
        {
            request.Content = new ByteArrayContent(content);
            request.Headers.Add("X-WOPI-Override", "PUT");
            request.Headers.Add("X-WOPI-Lock", "wopi-lock");
            using var put = await host.Client.SendAsync(request);
            await Assert.That(put.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        using var unlocked = await host.Post(fileId, token, "UNLOCK", ("X-WOPI-Lock", "wopi-lock"));
        await Assert.That(unlocked.StatusCode).IsEqualTo(HttpStatusCode.OK);
    }

    /// <summary>
    /// A file deleted through WOPI takes its desktop coauthoring state with it; a store that fails to forget it does not
    /// fail the delete.
    /// </summary>
    [Test]
    public async Task DeleteFile()
    {
        var (host, fileId, alice, _) = await Start();
        await using var _ = host;
        await alice.Send(alice.Join());
        await alice.Sync();
        await alice.Coauth("ExitCoauthoring");
        var store = host.Service<IFsshttpStore>();
        await Assert.That(await store.Get(fileId, FsshttpCells.ContentPartition)).IsNotNull();

        var token = await host.Token(TestServerHost.Access(fileId));
        using var response = await host.Post(fileId, token, "DELETE");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(await store.Get(fileId, FsshttpCells.ContentPartition)).IsNull();
        await Assert.That(await store.Get(fileId, CoauthSessions.SessionPartition)).IsNull();
    }

    [Test]
    public async Task DeleteFile_StoreFails()
    {
        var (host, fileId, alice, _) = await Start(services: _ => _.AddSingleton<IFsshttpStore, UndeletableStore>());
        await using var _ = host;
        await alice.Sync();
        var token = await host.Token(TestServerHost.Access(fileId));
        using var response = await host.Post(fileId, token, "DELETE");
        await Assert.That(response.StatusCode).IsEqualTo(HttpStatusCode.OK);
        await Assert.That(File.Exists(host.PathOf("doc.docx"))).IsFalse();
        await Assert.That(host.Log.Lines.Any(_ => _.StartsWith("Warning") && _.Contains("Deleting the desktop coauthoring state"))).IsTrue();
    }

    sealed class UndeletableStore :
        IFsshttpStore
    {
        MemoryFsshttpStore inner = new();

        public Task<FsshttpState?> Get(string fileId, Guid partition, Cancel cancel = default) =>
            inner.Get(fileId, partition, cancel);

        public Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FsshttpState value, Cancel cancel = default) =>
            inner.CompareExchange(fileId, partition, comparand, value, cancel);

        public Task Delete(string fileId, Cancel cancel = default) =>
            throw new IOException("Undeletable");
    }
}
