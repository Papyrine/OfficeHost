using System.Xml.Linq;

/// <summary>
/// Desktop Excel, as it went against SharePoint Server and the sample (docs/fsshttp-notes.md): it opens without a lock,
/// takes an exclusive lock at the first edit, with a LockStatus that runs only if the lock fails, saves with the lock
/// id as BypassLockID alone, and releases the lock on close.
/// </summary>
public class ExcelTests
{
    const string lockId = "{3D80B2EF-FA28-402D-A400-1926062384B0}";

    static byte[] Book(string text) => Zips.Create(("xl/worksheets/sheet1.xml", Zips.Text(text)), ("xl/styles.xml", Zips.Random(10_000, 4)));

    static async Task<(TestServerHost Host, string FileId, WordClient Alice, WordClient Bob)> Start(Mode mode = Mode.Edit)
    {
        var host = await TestServerHost.Start();
        await File.WriteAllBytesAsync(host.PathOf("book.xlsx"), Book("start"));
        var (fileId, token) = await host.Open("book.xlsx", mode);
        var bobToken = await host.Token(TestServerHost.Access(fileId, userId: "bob"));
        return (host, fileId, new(host, token), new(host, bobToken));
    }

    /// <summary>
    /// The exclusive lock at the first edit, with the LockStatus Excel sends with it.
    /// </summary>
    static Task<List<SubResult>> Lock(WordClient client, string type = "GetLock", string id = lockId) =>
        client.Send(
            WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", type), new XAttribute("ExclusiveLockID", id), new XAttribute("Timeout", 600))),
            WordClient.SubRequest("LockStatus", "2", null, "1", "OnFail"));

    static async Task<Dictionary<string, string>> DocProps(WordClient client)
    {
        var result = (await client.Send(WordClient.SubRequest("GetDocMetaInfo", "1"))).Single();
        return result.Data.Descendants()
            .Where(_ => _.Name.LocalName == "Property")
            .ToDictionary(_ => (string) _.Attribute("Key")!, _ => (string) _.Attribute("Value")!);
    }

    [Test]
    public async Task EditAndSave()
    {
        var (host, fileId, alice, _) = await Start();
        await using var _ = host;

        // the open: no lock, and a file Excel may edit
        var props = await DocProps(alice);
        await Assert.That(props["vti_canmaybeedit"]).IsEqualTo("true");
        await Assert.That(props.ContainsKey("vti_sourcecontrollockid")).IsFalse();
        await alice.Sync();

        // the first edit
        var locked = await Lock(alice);
        await Assert.That(locked.Select(_ => _.ErrorCode).ToList()).IsEquivalentTo(["Success", "DependentOnlyOnFailRequestSucceeded"]);
        await Assert.That((await host.Service<IOfficeLockStore>().Get(fileId))!.Id).IsEqualTo($"fsshttp-exclusive:{lockId}");
        props = await DocProps(alice);
        await Assert.That(props["vti_sourcecontrollockid"]).IsEqualTo(lockId.Trim('{', '}').ToLowerInvariant());
        await Assert.That(props["vti_sourcecontrollocktype"]).IsEqualTo("0");
        await Assert.That(props["vti_sourcecontrolcheckedoutby"]).IsEqualTo("Alice");

        // the save, with the lock id as BypassLockID alone
        await Assert.That(await alice.Save(Book("saved"), lockId, exclusive: true)).IsTypeOf<PutChangesResponse>();
        await Assert.That((await File.ReadAllBytesAsync(host.PathOf("book.xlsx"))).SameAs(Book("saved"))).IsTrue();

        // the close
        var released = await alice.Send(WordClient.SubRequest("ExclusiveLock", "1", WordClient.Data(new XAttribute("ExclusiveLockRequestType", "ReleaseLock"), new XAttribute("ExclusiveLockID", lockId))));
        await Assert.That(released.Single().ErrorCode).IsEqualTo("Success");
        await Assert.That(await host.Service<IOfficeLockStore>().Get(fileId)).IsNull();
    }

    /// <summary>
    /// A second Excel cannot take the lock or save. The refusal names the user in its ErrorMessage, which Excel shows in
    /// its File In Use, and the LockStatus has the lock, as SharePoint answers.
    /// </summary>
    [Test]
    public async Task LockedByAnother()
    {
        var (host, _, alice, bob) = await Start();
        await using var _ = host;
        await Lock(alice);
        await bob.Sync();

        var refused = await Lock(bob, id: "{B0B0B0B0-0000-4000-8000-000000000000}");
        await Assert.That(refused[0].ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That(refused[0].ErrorMessage).IsEqualTo("Alice");
        await Assert.That(refused[1].ErrorCode).IsEqualTo("Success");
        await Assert.That(refused[1].ErrorMessage).IsNull();
        await Assert.That(refused[1].Attribute("LockType")).IsEqualTo("2");
        await Assert.That(refused[1].Attribute("LockID")).IsEqualTo(lockId.Trim('{', '}').ToLowerInvariant());
        await Assert.That(refused[1].Attribute("LockedBy")).IsEqualTo("Alice");

        var save = await bob.Save(Book("bob"), "{B0B0B0B0-0000-4000-8000-000000000000}", exclusive: true);
        await Assert.That(((FailedSubResponse) save).Error.SupplementalInfo).IsEqualTo("FileAlreadyLockedOnServer");
        // the lock again, as Excel refreshes it
        await Assert.That((await Lock(alice, "RefreshLock"))[0].ErrorCode).IsEqualTo("Success");
    }

    /// <summary>
    /// The LockStatus of a file Word coauthors, of a file a WOPI client locked, and of a file nobody locked.
    /// </summary>
    [Test]
    public async Task LockStatus()
    {
        var (host, fileId, alice, bob) = await Start();
        await using var _ = host;

        static async Task<SubResult> Status(WordClient client) =>
            (await client.Send(WordClient.SubRequest("LockStatus", "1"))).Single();

        await Assert.That((await Status(bob)).Attribute("LockType")).IsEqualTo("0");

        await alice.Send(alice.Join());
        var schema = await Status(bob);
        await Assert.That(schema.Attribute("LockType")).IsEqualTo("1");
        await Assert.That(schema.Attribute("LockID")).IsEqualTo(WordClient.SchemaLockId.Trim('{', '}').ToLowerInvariant());
        await Assert.That(schema.Attribute("LockedBy")).IsEqualTo("Alice");
        var refused = (await Lock(bob))[0];
        await Assert.That(refused.ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        await Assert.That(refused.ErrorMessage).IsEqualTo("Alice");
        await alice.Coauth("ExitCoauthoring");

        var token = await host.Token(TestServerHost.Access(fileId, userId: "carol"));
        using (var wopi = await host.Post(fileId, token, "LOCK", ("X-WOPI-Lock", "wopi-lock")))
        {
            await Assert.That(wopi.StatusCode).IsEqualTo(HttpStatusCode.OK);
        }

        var wopiLock = await Status(bob);
        await Assert.That(wopiLock.Attribute("LockType")).IsEqualTo("2");
        await Assert.That(wopiLock.Attribute("LockID")).IsNull();
        var wopiRefused = (await Lock(bob))[0];
        await Assert.That(wopiRefused.ErrorCode).IsEqualTo("FileAlreadyLockedOnServer");
        // nobody to name for a WOPI lock
        await Assert.That(wopiRefused.ErrorMessage).IsNull();
    }

    /// <summary>
    /// A user who cannot write gets a file Excel may not edit, and no lock.
    /// </summary>
    [Test]
    public async Task ReadOnly()
    {
        var (host, _, alice, _) = await Start(Mode.ReadOnly);
        await using var _ = host;
        await Assert.That((await DocProps(alice))["vti_canmaybeedit"]).IsEqualTo("false");
        await Assert.That((await Lock(alice))[0].ErrorCode).IsEqualTo("FileUnauthorizedAccess");
    }
}
