/// <summary>
/// <see cref="CellState"/>, and <see cref="FsshttpCells"/>, which keeps the content partition of a file in step with the
/// file.
/// </summary>
public class FsshttpCellsTests : IDisposable
{
    TempDirectory folder = new();
    FileSystemStorage storage;
    MemoryFsshttpStore store = new();

    public FsshttpCellsTests() =>
        storage = new(folder.Path);

    public void Dispose() =>
        folder.Dispose();

    string Create(byte[] content)
    {
        File.WriteAllBytes(Path.Combine(folder, "file.docx"), content);
        return storage.GetFileId("file.docx");
    }

    static byte[] Zip(string text) => Zips.Create(("a.txt", Zips.Text(text)), ("b.bin", Zips.Random(10_000, 1)));

    [Test]
    public async Task CellState_RoundTrip()
    {
        var ids = new IdAllocator(Guid.NewGuid(), 5);
        var update = FileCell.Write(Zip("a"), null, ids, ChunkingOptions.Default);
        var state = new CellState(update.Storage, ids.Guid, ids.Next, "version");
        var read = CellState.Read(state.ToArray());
        await Assert.That(read.IdGuid).IsEqualTo(ids.Guid);
        await Assert.That(read.NextId).IsEqualTo(ids.Next);
        await Assert.That(read.FileVersion).IsEqualTo("version");
        await Assert.That(read.Storage!.StorageIndexId).IsEqualTo(update.Storage.StorageIndexId);
        await Assert.That(Fsshttpb.Describe(read.Storage.Elements.Values.OrderBy(_ => _.Id.ToString())))
            .IsEqualTo(Fsshttpb.Describe(update.Storage.Elements.Values.OrderBy(_ => _.Id.ToString())));
        await Assert.That(FileCell.Read(read.Storage).Root.ToArray().AsSpan().SequenceEqual(Zip("a"))).IsTrue();

        var empty = CellState.New();
        var readEmpty = CellState.Read(empty.ToArray());
        await Assert.That(readEmpty.Storage).IsNull();
        await Assert.That(readEmpty.FileVersion).IsNull();
        await Assert.That(readEmpty.NextId).IsEqualTo(1ul);
        await Assert.That(readEmpty.Ids().Guid).IsEqualTo(empty.IdGuid);
    }

    [Test]
    [Arguments("")]
    [Arguments("4F484331")]
    [Arguments("58585858")]
    [Arguments("4F48433102")]
    [Arguments("4F4843310000000000000000000000000000000000000300")]
    public async Task CellState_Corrupt(string hex)
    {
        var bytes = Fsshttpb.Hex(hex);
        await Assert.That(() => CellState.Read(bytes)).Throws<InvalidDataException>();
    }

    /// <summary>
    /// The content is written from the file on first use, then reused while the file keeps its version.
    /// </summary>
    [Test]
    public async Task GetContent_Unchanged()
    {
        var id = Create(Zip("a"));
        var cells = new FsshttpCells(store, storage);
        var first = (await cells.GetContent(id))!;
        await Assert.That(first.Version).IsNotNull();
        await Assert.That(first.State.FileVersion).IsEqualTo((await storage.GetFile(id))!.Version);
        await Assert.That(FileCell.Read(first.State.Storage!).Root.ToArray().AsSpan().SequenceEqual(Zip("a"))).IsTrue();

        var second = (await cells.GetContent(id))!;
        await Assert.That(second.Version).IsEqualTo(first.Version);
    }

    /// <summary>
    /// A file changed outside FSSHTTP is written as the next revision: the chunks that did not change keep their
    /// objects, and the Serial Numbers keep increasing.
    /// </summary>
    [Test]
    public async Task GetContent_FileChanged()
    {
        var id = Create(Zip("a"));
        var cells = new FsshttpCells(store, storage);
        var first = (await cells.GetContent(id))!;
        await storage.Write(id, new MemoryStream(Zip("b")));
        var second = (await cells.GetContent(id))!;
        await Assert.That(second.Version)!.IsNotEqualTo(first.Version);
        await Assert.That(FileCell.Read(second.State.Storage!).Root.ToArray().AsSpan().SequenceEqual(Zip("b"))).IsTrue();
        await Assert.That(second.State.IdGuid).IsEqualTo(first.State.IdGuid);
        await Assert.That(second.State.NextId).IsGreaterThan(first.State.NextId);
        var kept = FileCell.Read(second.State.Storage!).Root.Children.Select(_ => _.Id)
            .Intersect(FileCell.Read(first.State.Storage!).Root.Children.Select(_ => _.Id));
        await Assert.That(kept.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task GetContent_NoFile()
    {
        var cells = new FsshttpCells(store, storage);
        await Assert.That(await cells.GetContent("missing")).IsNull();
    }

    /// <summary>
    /// A state that is not a file stream is replaced by one written from nothing.
    /// </summary>
    [Test]
    public async Task GetContent_BrokenState()
    {
        var id = Create(Zip("a"));
        var ids = new IdAllocator(Guid.NewGuid());
        var broken = CellStorage.Create([], ids.NewId());
        var state = new CellState(broken, ids.Guid, ids.Next, "old");
        var cells = new FsshttpCells(store, storage);
        await cells.TryUpdate(id, FsshttpCells.ContentPartition, null, state);
        var content = (await cells.GetContent(id))!;
        await Assert.That(FileCell.Read(content.State.Storage!).Root.ToArray().AsSpan().SequenceEqual(Zip("a"))).IsTrue();
    }

    /// <summary>
    /// A state written by another instance between the read and the write: the write fails, and is done again from it.
    /// </summary>
    [Test]
    public async Task GetContent_Race()
    {
        var id = Create(Zip("a"));
        var racing = new RacingStore(store, 1);
        var cells = new FsshttpCells(racing, storage);
        var content = (await cells.GetContent(id))!;
        await Assert.That(racing.Lost).IsEqualTo(1);
        await Assert.That(FileCell.Read(content.State.Storage!).Root.ToArray().AsSpan().SequenceEqual(Zip("a"))).IsTrue();

        var always = new FsshttpCells(new RacingStore(new MemoryFsshttpStore(), int.MaxValue), storage);
        await Assert.That(() => always.GetContent(id)).Throws<InvalidOperationException>();
    }

    /// <summary>
    /// A store where another writer comes first, a given number of times.
    /// </summary>
    class RacingStore(IFsshttpStore inner, int times) :
        IFsshttpStore
    {
        public int Lost { get; private set; }

        public Task<FsshttpState?> Get(string fileId, Guid partition, Cancel cancel = default) =>
            inner.Get(fileId, partition, cancel);

        public async Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FsshttpState value, Cancel cancel = default)
        {
            if (Lost >= times)
            {
                return await inner.CompareExchange(fileId, partition, comparand, value, cancel);
            }

            Lost++;
            // another instance writes the same file meanwhile, with a state of an older version of the file
            var ids = new IdAllocator(Guid.NewGuid());
            var other = FileCell.Write(Zips.Text("other"), null, ids, ChunkingOptions.Default);
            await inner.CompareExchange(fileId, partition, comparand, new(Guid.NewGuid().ToString(), new CellState(other.Storage, ids.Guid, ids.Next, "older").ToArray()), cancel);
            return false;
        }

        public Task Delete(string fileId, Cancel cancel = default) =>
            inner.Delete(fileId, cancel);
    }

    [Test]
    public async Task Get_Partitions()
    {
        var cells = new FsshttpCells(store, storage);
        var metadata = await cells.Get("file", FsshttpCells.MetadataPartition);
        await Assert.That(metadata.Version).IsNull();
        await Assert.That(metadata.State.Storage).IsNull();
        var stored = (await cells.TryUpdate("file", FsshttpCells.MetadataPartition, null, metadata.State))!;
        await Assert.That(await cells.TryUpdate("file", FsshttpCells.MetadataPartition, null, metadata.State)).IsNull();
        await Assert.That((await cells.Get("file", FsshttpCells.MetadataPartition)).Version).IsEqualTo(stored.Version);
        await Assert.That((await cells.Get("file", FsshttpCells.EditorsTablePartition)).Version).IsNull();
        await Assert.That(FsshttpCells.ContentPartition).IsEqualTo(Guid.Empty);
    }
}
