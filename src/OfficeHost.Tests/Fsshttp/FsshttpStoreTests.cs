/// <summary>
/// The contract of <see cref="IFsshttpStore"/>, for <see cref="MemoryFsshttpStore"/> and
/// <see cref="FileSystemFsshttpStore"/>.
/// </summary>
public class FsshttpStoreTests : IDisposable
{
    TempDirectory folder = new();

    public void Dispose() =>
        folder.Dispose();

    public static IEnumerable<string> Stores() => ["Memory", "FileSystem"];

    IFsshttpStore Store(string name)
    {
        if (name == "Memory")
        {
            return new MemoryFsshttpStore();
        }

        return new FileSystemFsshttpStore(Path.Combine(folder, "states"));
    }

    static readonly Guid partition = Guid.NewGuid();
    static readonly Guid other = Guid.NewGuid();

    static FsshttpState State(string version, params byte[] data) => new(version, data);

    [Test]
    [MethodDataSource(nameof(Stores))]
    public async Task CompareExchange(string name)
    {
        var store = Store(name);
        await Assert.That(await store.Get("file", partition)).IsNull();
        // a state that is not there yet is replaced from null only
        await Assert.That(await store.CompareExchange("file", partition, "1", State("2", 1))).IsFalse();
        await Assert.That(await store.CompareExchange("file", partition, null, State("1", 1, 2, 3))).IsTrue();
        var state = (await store.Get("file", partition))!;
        await Assert.That(state.Version).IsEqualTo("1");
        await Assert.That(state.Data.ToArray()).IsEquivalentTo(new byte[] { 1, 2, 3 });

        await Assert.That(await store.CompareExchange("file", partition, null, State("x"))).IsFalse();
        await Assert.That(await store.CompareExchange("file", partition, "0", State("x"))).IsFalse();
        await Assert.That(await store.CompareExchange("file", partition, "1", State("2", 4))).IsTrue();
        await Assert.That((await store.Get("file", partition))!.Version).IsEqualTo("2");
        await Assert.That((await store.Get("file", partition))!.Data.ToArray()).IsEquivalentTo(new byte[] { 4 });
    }

    /// <summary>
    /// Partitions and files are apart; deleting a file forgets all its partitions, and only them.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Stores))]
    public async Task Partitions(string name)
    {
        var store = Store(name);
        await store.CompareExchange("file", partition, null, State("a", 1));
        await store.CompareExchange("file", other, null, State("b", 2));
        await store.CompareExchange("another/file é", partition, null, State("c", 3));
        await Assert.That((await store.Get("file", other))!.Version).IsEqualTo("b");
        await Assert.That((await store.Get("another/file é", partition))!.Version).IsEqualTo("c");

        await store.Delete("file");
        await Assert.That(await store.Get("file", partition)).IsNull();
        await Assert.That(await store.Get("file", other)).IsNull();
        await Assert.That((await store.Get("another/file é", partition))!.Version).IsEqualTo("c");
        await store.Delete("missing");
    }

    /// <summary>
    /// Of concurrent writers from the same version, one wins.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(Stores))]
    public async Task Concurrent(string name)
    {
        var store = Store(name);
        await store.CompareExchange("file", partition, null, State("0"));
        var results = await Task.WhenAll(Enumerable.Range(1, 20).Select(_ => Task.Run(() => store.CompareExchange("file", partition, "0", State($"{_}")))));
        await Assert.That(results.Count(_ => _)).IsEqualTo(1);
    }

    /// <summary>
    /// The states in a folder outlive the store, and an empty state and a long version are kept.
    /// </summary>
    [Test]
    public async Task FileSystem_Reopened()
    {
        var root = Path.Combine(folder, "states");
        var store = new FileSystemFsshttpStore(root);
        var version = new string('v', 1000);
        await store.CompareExchange("file", partition, null, State(version));
        var reopened = new FileSystemFsshttpStore(root);
        var state = (await reopened.Get("file", partition))!;
        await Assert.That(state.Version).IsEqualTo(version);
        await Assert.That(state.Data.Length).IsEqualTo(0);
        await Assert.That(reopened.Root).IsEqualTo(Path.GetFullPath(root));
        // no temporary file is left
        await Assert.That(Directory.GetFiles(root, "*.tmp", SearchOption.AllDirectories).Length).IsEqualTo(0);
    }

    [Test]
    [Arguments(new byte[] { 1, 2 })]
    [Arguments(new byte[] { 10, 0, 0, 0, 65 })]
    [Arguments(new byte[] { 255, 255, 255, 255 })]
    public async Task FileSystem_Truncated(byte[] content)
    {
        var root = Path.Combine(folder, "states");
        var store = new FileSystemFsshttpStore(root);
        await store.CompareExchange("file", partition, null, State("1"));
        var path = Directory.GetFiles(root, "*.state", SearchOption.AllDirectories).Single();
        await File.WriteAllBytesAsync(path, content);
        await Assert.That(() => store.Get("file", partition)).Throws<InvalidDataException>();
    }
}
