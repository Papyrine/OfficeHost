/// <summary>
/// <see cref="FileSystemStorage"/> and <see cref="MemoryLockStore"/>.
/// </summary>
public class StorageTests : IDisposable
{
    readonly TempFolder folder = new();
    readonly FileSystemStorage storage;

    public StorageTests() =>
        storage = new(folder.Path);

    public void Dispose() =>
        folder.Dispose();

    string Create(string path, string content = "content")
    {
        var full = folder.Combine(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);
        return storage.GetFileId(path);
    }

    static async Task<string> Read(Stream stream)
    {
        await using (stream)
        {
            using var reader = new StreamReader(stream);
            return await reader.ReadToEndAsync();
        }
    }

    [Test]
    public async Task GetFile()
    {
        var id = Create("sub/test.docx", "12345");
        var file = (await storage.GetFile(id))!;
        await Assert.That(file.Id).IsEqualTo(id);
        await Assert.That(file.Name).IsEqualTo("test.docx");
        await Assert.That(file.Size).IsEqualTo(5);
        await Assert.That(file.OwnerId).IsEqualTo("owner");
        await Assert.That(file.LastModified).IsEqualTo(new DateTimeOffset(File.GetLastWriteTimeUtc(folder.Combine("sub/test.docx"))));
        await Assert.That(storage.GetPath(id)).IsEqualTo(folder.Combine("sub", "test.docx"));
    }

    [Test]
    public async Task GetFile_Unknown()
    {
        await Assert.That(await storage.GetFile("unknown")).IsNull();
        var id = Create("test.docx");
        File.Delete(folder.Combine("test.docx"));
        await Assert.That(await storage.GetFile(id)).IsNull();
    }

    /// <summary>
    /// Ids are URL safe, given once per file, and kept across instances of the storage.
    /// </summary>
    [Test]
    public async Task FileIds()
    {
        var a = Create("a.docx");
        var b = Create("b.docx");
        await Assert.That(a).IsNotEqualTo(b);
        await Assert.That(storage.GetFileId("a.docx")).IsEqualTo(a);
        await Assert.That(storage.GetFileId("/a.docx")).IsEqualTo(a);
        await Assert.That(a.All(_ => char.IsAsciiLetterOrDigit(_) || _ is '-' or '_')).IsTrue();
        var reopened = new FileSystemStorage(folder.Path);
        await Assert.That(reopened.GetFileId("b.docx")).IsEqualTo(b);
        await Assert.That((await reopened.GetFile(a))!.Name).IsEqualTo("a.docx");
    }

    [Test]
    public async Task FileIds_OnlyForFilesOfTheStorage()
    {
        await Helpers.Throws<FileNotFoundException>(() => Task.FromResult(storage.GetFileId("missing.docx")));
        Directory.CreateDirectory(folder.Combine("folder"));
        await Helpers.Throws<FileNotFoundException>(() => Task.FromResult(storage.GetFileId("folder")));
        Create("a.docx");
        foreach (var path in new[] { "../outside.docx", ".wopi/ids.json", "" })
        {
            await Helpers.Throws<UnauthorizedAccessException>(() => Task.FromResult(storage.GetFileId(path)));
        }
    }

    [Test]
    public async Task OpenRead()
    {
        var id = Create("test.docx", "content");
        await Assert.That(await Read(await storage.OpenRead(id))).IsEqualTo("content");
        File.Delete(folder.Combine("test.docx"));
        await Helpers.Throws<FileNotFoundException>(() => storage.OpenRead(id));
        await Helpers.Throws<FileNotFoundException>(() => storage.OpenRead("unknown"));
    }

    /// <summary>
    /// A write replaces the content atomically, so a reader keeps the content it opened.
    /// </summary>
    [Test]
    public async Task Write()
    {
        var id = Create("test.docx", "before");
        var reader = await storage.OpenRead(id);
        var written = await storage.Write(id, Helpers.Stream("after"));
        await Assert.That(written.Size).IsEqualTo(5);
        await Assert.That(await Read(reader)).IsEqualTo("before");
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("test.docx"))).IsEqualTo("after");
        await Assert.That(Directory.GetFiles(folder.Combine(".wopi", "tmp"))).IsEmpty();
    }

    /// <summary>
    /// Every write is a new version, even within the resolution of the clock of the file system.
    /// </summary>
    [Test]
    public async Task Write_NewVersions()
    {
        var id = Create("test.docx");
        var versions = new List<long>
        {
            long.Parse((await storage.GetFile(id))!.Version)
        };
        for (var i = 0; i < 5; i++)
        {
            versions.Add(long.Parse((await storage.Write(id, Helpers.Stream("same"))).Version));
        }

        for (var i = 1; i < versions.Count; i++)
        {
            await Assert.That(versions[i]).IsGreaterThan(versions[i - 1]);
        }
    }

    [Test]
    public async Task Write_Unknown()
    {
        var id = Create("test.docx");
        File.Delete(folder.Combine("test.docx"));
        await Helpers.Throws<FileNotFoundException>(() => storage.Write(id, Helpers.Stream("x")));
        await Assert.That(File.Exists(folder.Combine("test.docx"))).IsFalse();
    }

    [Test]
    public async Task GetSibling()
    {
        var id = Create("sub/test.docx");
        var otherId = Create("sub/other.docx");
        Create("other.docx");
        await Assert.That((await storage.GetSibling(id, "other.docx"))!.Id).IsEqualTo(otherId);
        await Assert.That(await storage.GetSibling(id, "missing.docx")).IsNull();
    }

    [Test]
    public async Task CreateSibling()
    {
        var id = Create("sub/test.docx");
        var created = (await storage.CreateSibling(id, "new.pdf", Helpers.Stream("pdf"), false))!;
        await Assert.That(created.Name).IsEqualTo("new.pdf");
        await Assert.That(created.Id).IsEqualTo(storage.GetFileId("sub/new.pdf"));
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("sub/new.pdf"))).IsEqualTo("pdf");

        await Assert.That(await storage.CreateSibling(id, "new.pdf", Helpers.Stream("again"), false)).IsNull();
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("sub/new.pdf"))).IsEqualTo("pdf");

        var overwritten = (await storage.CreateSibling(id, "new.pdf", Helpers.Stream("again"), true))!;
        await Assert.That(overwritten.Id).IsEqualTo(created.Id);
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("sub/new.pdf"))).IsEqualTo("again");
    }

    /// <summary>
    /// The storage refuses the names that are not in the folder of the file, should a caller not check them.
    /// </summary>
    [Test]
    [Arguments("../escape.docx")]
    [Arguments("deeper/new.docx")]
    [Arguments("..")]
    public async Task Siblings_OnlyInTheFolder(string name)
    {
        var id = Create("sub/test.docx");
        await Helpers.Throws<UnauthorizedAccessException>(() => storage.CreateSibling(id, name, Helpers.Stream("x"), true));
        await Helpers.Throws<UnauthorizedAccessException>(() => storage.GetSibling(id, name));
        await Helpers.Throws<UnauthorizedAccessException>(() => storage.Rename(id, name));
    }

    /// <summary>
    /// A renamed file keeps its id, and the id of the old name is no more.
    /// </summary>
    [Test]
    public async Task Rename()
    {
        var id = Create("test.docx", "content");
        var renamed = (await storage.Rename(id, "renamed.docx"))!;
        await Assert.That(renamed.Id).IsEqualTo(id);
        await Assert.That(renamed.Name).IsEqualTo("renamed.docx");
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("renamed.docx"))).IsEqualTo("content");
        await Assert.That(storage.GetFileId("renamed.docx")).IsEqualTo(id);
        await Assert.That((await new FileSystemStorage(folder.Path).GetFile(id))!.Name).IsEqualTo("renamed.docx");
        await File.WriteAllTextAsync(folder.Combine("test.docx"), "new file");
        await Assert.That(storage.GetFileId("test.docx")).IsNotEqualTo(id);
    }

    [Test]
    public async Task Rename_ExistingName()
    {
        var id = Create("test.docx", "content");
        Create("other.docx", "other");
        await Assert.That(await storage.Rename(id, "other.docx")).IsNull();
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("other.docx"))).IsEqualTo("other");
        await Assert.That(await File.ReadAllTextAsync(folder.Combine("test.docx"))).IsEqualTo("content");
    }

    [Test]
    public async Task Rename_CaseOnly()
    {
        var id = Create("test.docx");
        var renamed = (await storage.Rename(id, "Test.docx"))!;
        await Assert.That(renamed.Name).IsEqualTo("Test.docx");
        await Assert.That(storage.GetFileId("Test.docx")).IsEqualTo(id);
    }

    [Test]
    public async Task Delete()
    {
        var id = Create("test.docx");
        await storage.Delete(id);
        await Assert.That(File.Exists(folder.Combine("test.docx"))).IsFalse();
        await Assert.That(await storage.GetFile(id)).IsNull();
        await Helpers.Throws<FileNotFoundException>(() => storage.Delete(id));
    }

    [Test]
    public async Task MissingRoot() =>
        await Helpers.Throws<DirectoryNotFoundException>(() => Task.FromResult(new FileSystemStorage(folder.Combine("missing"))));

    // MemoryLockStore

    [Test]
    public async Task Locks_CompareExchange()
    {
        var time = new ManualTime();
        var locks = new MemoryLockStore(time);
        var expires = time.Now.AddMinutes(30);
        await Assert.That(await locks.Get("f")).IsNull();

        // lock an unlocked file
        await Assert.That(await locks.CompareExchange("f", null, new("L", expires))).IsNull();
        await Assert.That((await locks.Get("f"))!.Id).IsEqualTo("L");

        // a mismatch changes nothing, and gives the current lock
        await Assert.That((await locks.CompareExchange("f", "other", new("X", expires)))!.Id).IsEqualTo("L");
        await Assert.That((await locks.CompareExchange("f", null, new("X", expires)))!.Id).IsEqualTo("L");
        await Assert.That((await locks.Get("f"))!.Id).IsEqualTo("L");

        // swap, then unlock
        await Assert.That((await locks.CompareExchange("f", "L", new("L2", expires)))!.Id).IsEqualTo("L");
        await Assert.That((await locks.CompareExchange("f", "L2", null))!.Id).IsEqualTo("L2");
        await Assert.That(await locks.Get("f")).IsNull();
    }

    [Test]
    public async Task Locks_Expire()
    {
        var time = new ManualTime();
        var locks = new MemoryLockStore(time);
        await locks.CompareExchange("f", null, new("L", time.Now.AddMinutes(30)));
        time.Advance(TimeSpan.FromMinutes(30));
        await Assert.That(await locks.Get("f")).IsNull();
        // an expired lock is no lock: it cannot be swapped, and does not prevent a new one
        await Assert.That(await locks.CompareExchange("f", "L", new("L2", time.Now.AddMinutes(30)))).IsNull();
        await Assert.That(await locks.Get("f")).IsNull();
        await Assert.That(await locks.CompareExchange("f", null, new("L3", time.Now.AddMinutes(30)))).IsNull();
        await Assert.That((await locks.Get("f"))!.Id).IsEqualTo("L3");
    }

    /// <summary>
    /// Of concurrent lock attempts on an unlocked file, exactly one wins.
    /// </summary>
    [Test]
    public async Task Locks_Race()
    {
        var locks = new MemoryLockStore(TimeProvider.System);
        var expires = DateTimeOffset.UtcNow.AddMinutes(30);
        var results = await Task.WhenAll(Enumerable.Range(0, 50).Select(_ => Task.Run(() => locks.CompareExchange("f", null, new($"L{_}", expires)))));
        await Assert.That(results.Count(_ => _ == null)).IsEqualTo(1);
    }
}
