/// <summary>
/// What every <see cref="IOfficeStorage"/> of the packages does, linked into the tests of each: a missing file throws,
/// names are unique in a folder ignoring case, and ids survive renames.
/// </summary>
public abstract class StorageContractTests
{
    // each test has a folder and files of its own
    protected string folder = Guid.NewGuid().ToString("N");

    protected abstract IOfficeStorage Storage { get; }

    /// <summary>
    /// Adds a file of the owner <c>owner</c>, or returns null when the folder has a file of that name.
    /// </summary>
    protected abstract Task<OfficeFile?> Add(string folder, string name, Stream content);

    protected static MemoryStream Content(string content) =>
        new(Encoding.UTF8.GetBytes(content));

    protected static async Task<string> Read(Stream stream)
    {
        await using (stream)
        {
            using var reader = new StreamReader(stream);
            return await reader.ReadToEndAsync();
        }
    }

    [Test]
    public async Task Storage_Lifecycle()
    {
        var storage = Storage;
        var file = (await Add(folder, "test.docx", Content("12345")))!;
        await Assert.That(file.Name).IsEqualTo("test.docx");
        await Assert.That(file.Size).IsEqualTo(5);
        await Assert.That(file.OwnerId).IsEqualTo("owner");
        await Assert.That(await storage.GetFile(file.Id)).IsEqualTo(file);
        await Assert.That(await Add(folder, "TEST.docx", Content(""))).IsNull();
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("12345");

        var written = await storage.Write(file.Id, Content("changed"));
        await Assert.That(written.Id).IsEqualTo(file.Id);
        await Assert.That(written.Size).IsEqualTo(7);
        await Assert.That(written.Version).IsNotEqualTo(file.Version);
        await Assert.That(await Read(await storage.OpenRead(file.Id))).IsEqualTo("changed");

        await Assert.That(await storage.GetSibling(file.Id, "other.docx")).IsNull();
        var sibling = (await storage.CreateSibling(file.Id, "other.docx", Content("other"), false))!;
        await Assert.That(sibling.Id).IsNotEqualTo(file.Id);
        await Assert.That(sibling.OwnerId).IsEqualTo("owner");
        await Assert.That(await storage.GetSibling(file.Id, "other.docx")).IsEqualTo(sibling);
        await Assert.That(await storage.CreateSibling(file.Id, "other.docx", Content("again"), false)).IsNull();
        var overwritten = (await storage.CreateSibling(file.Id, "other.docx", Content("again"), true))!;
        await Assert.That(overwritten.Id).IsEqualTo(sibling.Id);
        await Assert.That(await Read(await storage.OpenRead(sibling.Id))).IsEqualTo("again");

        // a sibling is in the same folder only
        var elsewhere = (await Add($"{folder}/sub", "other.docx", Content("")))!;
        await Assert.That(await storage.GetSibling(elsewhere.Id, "test.docx")).IsNull();

        await Assert.That(await storage.Rename(file.Id, "other.docx")).IsNull();
        await Assert.That((await storage.GetFile(file.Id))!.Name).IsEqualTo("test.docx");
        var renamed = (await storage.Rename(file.Id, "renamed.docx"))!;
        await Assert.That(renamed.Id).IsEqualTo(file.Id);
        await Assert.That(renamed.Name).IsEqualTo("renamed.docx");
        await Assert.That(renamed.Version).IsEqualTo(written.Version);
        await Assert.That((await storage.GetSibling(sibling.Id, "renamed.docx"))!.Id).IsEqualTo(file.Id);
        // the old name is free
        await Assert.That(await storage.GetSibling(sibling.Id, "test.docx")).IsNull();
        await Assert.That(await Add(folder, "test.docx", Content(""))).IsNotNull();
        // a change of case only is the same file
        await Assert.That((await storage.Rename(file.Id, "Renamed.docx"))!.Name).IsEqualTo("Renamed.docx");

        await storage.Delete(file.Id);
        await Assert.That(await storage.GetFile(file.Id)).IsNull();
        await Assert.That(await Add(folder, "renamed.docx", Content(""))).IsNotNull();
    }

    [Test]
    public async Task Storage_Missing()
    {
        var storage = Storage;
        await Assert.That(await storage.GetFile("missing")).IsNull();
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.OpenRead("missing"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Write("missing", Content("")));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.GetSibling("missing", "a.docx"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.CreateSibling("missing", "a.docx", Content(""), true));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Rename("missing", "a.docx"));
        await Assert.ThrowsAsync<FileNotFoundException>(() => storage.Delete("missing"));
    }

    [Test]
    public async Task Storage_Large()
    {
        var bytes = RandomNumberGenerator.GetBytes(20 * 1024 * 1024);
        var file = (await Add(folder, "large.docx", new MemoryStream(bytes)))!;
        await Assert.That(file.Size).IsEqualTo(bytes.Length);
        await using var stream = await Storage.OpenRead(file.Id);
        var read = new MemoryStream();
        await stream.CopyToAsync(read);
        await Assert.That(read.ToArray().SequenceEqual(bytes)).IsTrue();
    }

    [Test]
    public async Task Storage_Unicode()
    {
        var file = (await Add($"{folder}/données", "résumé ü 日本.docx", Content("")))!;
        await Assert.That((await Storage.GetFile(file.Id))!.Name).IsEqualTo("résumé ü 日本.docx");
        await Assert.That((await Storage.GetSibling(file.Id, "RÉSUMÉ Ü 日本.docx"))!.Id).IsEqualTo(file.Id);
    }

    /// <summary>
    /// Of siblings created at once with the same name, without overwriting, one is.
    /// </summary>
    [Test]
    public async Task Storage_CreateSibling_Race()
    {
        var source = (await Add(folder, "source.docx", Content("")))!;
        var results = await Task.WhenAll(
            Enumerable.Range(0, 10)
                .Select(_ => Task.Run(() => Storage.CreateSibling(source.Id, "race.docx", Content($"{_}"), false))));
        await Assert.That(results.Count(_ => _ != null)).IsEqualTo(1);
    }

    /// <summary>
    /// Of renames at once to the same name, one is.
    /// </summary>
    [Test]
    public async Task Storage_Rename_Race()
    {
        var files = new List<OfficeFile>();
        for (var index = 0; index < 10; index++)
        {
            files.Add((await Add(folder, $"{index}.docx", Content("")))!);
        }

        var results = await Task.WhenAll(files.Select(file => Task.Run(() => Storage.Rename(file.Id, "race.docx"))));
        await Assert.That(results.Count(_ => _ != null)).IsEqualTo(1);
    }
}
