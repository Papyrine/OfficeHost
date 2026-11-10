/// <summary>
/// The data elements of the parts of a save waiting for its last part (<see cref="PartialPuts"/>): a last part takes
/// those of its save only, those of a save that never ended do not stay in the store, and what waits for a file is
/// bounded.
/// </summary>
public class PartialPutsTests
{
    static readonly Guid guid = FssHttpB.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");

    static ObjectDataBlob Blob(uint id, int length) =>
        new(new(guid, id), new(guid, id), new byte[length]);

    /// <summary>
    /// The cell storage a last part made: what it holds of the parts is what its save took.
    /// </summary>
    static CellStorage Holding(params DataElement[] elements) =>
        CellStorage.Create(elements, new(guid, 1000));

    static (PartialPuts Puts, MemoryFssHttpStore Store, ManualTime Time) Start()
    {
        var store = new MemoryFssHttpStore();
        var time = new ManualTime();
        return (new(store, time), store, time);
    }

    /// <summary>
    /// What the store has for the file: the list of the parts, and the data elements of each in its slot.
    /// </summary>
    static async Task<int> Stored(MemoryFssHttpStore store)
    {
        var length = (await store.Get("file", PartialPuts.Partition))?.Data.Length ?? 0;
        for (var slot = 0; slot < 8; slot++)
        {
            length += Math.Max(0, await Slot(store, slot));
        }

        return length;
    }

    /// <summary>
    /// The size of what a slot has, -1 when it was never taken.
    /// </summary>
    static async Task<int> Slot(MemoryFssHttpStore store, int slot) =>
        (await store.Get("file", PartialPuts.PartitionOf(slot)))?.Data.Length ?? -1;

    /// <summary>
    /// The parts of a save that never ended (Word closed, or lost its connection) are not given to a last part after an
    /// hour, and leave the store with the next Put Changes of the file that applies, of any user; those still waiting
    /// stay.
    /// </summary>
    [Test]
    public async Task Expired_LeaveWithTheNextPut()
    {
        var (puts, store, time) = Start();
        var partition = FssHttpCells.ContentPartition;
        await puts.Add("file", "alice", partition, [Blob(1, 50_000)]);
        time.Advance(TimeSpan.FromMinutes(50));
        await puts.Add("file", "bob", partition, [Blob(2, 10_000)]);
        await Assert.That(await Stored(store)).IsGreaterThan(60_000);

        // a save of carol's: nothing has expired, nothing leaves
        await puts.Clean("file", "carol", partition, null);
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(1);
        await Assert.That(await Stored(store)).IsGreaterThan(60_000);

        // alice's have waited more than an hour
        time.Advance(TimeSpan.FromMinutes(20));
        await Assert.That(await puts.Get("file", "alice", partition)).IsEmpty();
        await Assert.That(await Stored(store)).IsGreaterThan(60_000);
        await puts.Clean("file", "carol", partition, null);
        await Assert.That(await Stored(store)).IsLessThan(20_000);
        await Assert.That(await puts.Get("file", "bob", partition)).Count().IsEqualTo(1);

        // bob's last part applied
        await puts.Clean("file", "bob", partition, Holding(Blob(2, 10_000)));
        await Assert.That(await puts.Get("file", "bob", partition)).IsEmpty();
        await Assert.That(await Stored(store)).IsLessThan(100);
    }

    /// <summary>
    /// Two saves of one user at once, as of Word on two machines: nothing in a part names its save, so the last part of
    /// one takes the data elements its save holds, and leaves those of the other for its own last part. Each part waits
    /// its own hour, whatever the user sends after it, and a part sent again replaces the one waiting.
    /// </summary>
    [Test]
    public async Task LastPart_TakesThoseOfItsSave()
    {
        var (puts, _, time) = Start();
        var partition = FssHttpCells.ContentPartition;
        await puts.Add("file", "alice", partition, [Blob(1, 1000), Blob(2, 1000)]);
        time.Advance(TimeSpan.FromMinutes(10));
        await puts.Add("file", "alice", partition, [Blob(3, 1000)]);
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(3);

        // the last part of the first save, which a last part of the metadata does not touch
        await puts.Clean("file", "alice", Guid.NewGuid(), Holding(Blob(1, 1000), Blob(2, 1000), Blob(3, 1000)));
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(3);
        await puts.Clean("file", "alice", partition, Holding(Blob(1, 1000), Blob(2, 1000)));
        var waiting = await puts.Get("file", "alice", partition);
        await Assert.That(waiting.Single().Id).IsEqualTo(Blob(3, 1000).Id);

        // another part of alice's does not keep the first waiting longer
        time.Advance(TimeSpan.FromMinutes(45));
        await puts.Add("file", "alice", partition, [Blob(4, 1000)]);
        await puts.Add("file", "alice", partition, [Blob(4, 2000)]);
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(2);
        time.Advance(TimeSpan.FromMinutes(20));
        waiting = await puts.Get("file", "alice", partition);
        await Assert.That(((ObjectDataBlob) waiting.Single()).Data.Length).IsEqualTo(2000);
    }

    /// <summary>
    /// Each part has an entry of its own in the store, a slot, next to the list of them: a part costs the store its own
    /// size, not that of the parts before it, and a Put Changes with none to forget reads the list alone. The store
    /// cannot forget a partition, so the slot of a part that is gone is emptied, and taken again by a later part. No
    /// slot is a cell a client may query or put.
    /// </summary>
    [Test]
    public async Task Parts_EachInASlot()
    {
        var store = new CountingStore();
        var puts = new PartialPuts(store, new ManualTime());
        var partition = FssHttpCells.ContentPartition;
        await puts.Add("file", "alice", partition, [Blob(1, 10_000)]);
        await puts.Add("file", "alice", partition, [Blob(2, 10_000)]);
        store.Written = 0;
        await puts.Add("file", "alice", partition, [Blob(3, 10_000)]);
        // the third part and the list, not the two before it
        await Assert.That(store.Written).IsGreaterThan(10_000).And.IsLessThan(11_000);
        await Assert.That(await Slot(store.Inner, 2)).IsGreaterThan(10_000);
        await Assert.That(await Slot(store.Inner, 3)).IsEqualTo(-1);

        // a Put Changes of bob's: the list alone
        store.Read = 0;
        await puts.Clean("file", "bob", partition, null);
        await Assert.That(store.Read).IsLessThan(1000);
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(3);

        // the last part of a save of the first two: their slots are emptied, and the next part takes the first
        await puts.Clean("file", "alice", partition, Holding(Blob(1, 10_000), Blob(2, 10_000)));
        await Assert.That(await Slot(store.Inner, 0)).IsEqualTo(0);
        await Assert.That(await Slot(store.Inner, 1)).IsEqualTo(0);
        await Assert.That(await Slot(store.Inner, 2)).IsGreaterThan(10_000);
        await puts.Add("file", "bob", partition, [Blob(4, 5000)]);
        await Assert.That(await Slot(store.Inner, 0)).IsGreaterThan(5000);
        await Assert.That(await Slot(store.Inner, 3)).IsEqualTo(-1);
        await Assert.That((await puts.Get("file", "alice", partition)).Single().Id).IsEqualTo(Blob(3, 0).Id);
        await Assert.That((await puts.Get("file", "bob", partition)).Single().Id).IsEqualTo(Blob(4, 0).Id);

        await Assert.That(PartialPuts.Owns(PartialPuts.Partition)).IsTrue();
        await Assert.That(PartialPuts.Owns(PartialPuts.PartitionOf(0))).IsTrue();
        await Assert.That(PartialPuts.Owns(PartialPuts.PartitionOf(123_456))).IsTrue();
        await Assert.That(PartialPuts.Owns(partition)).IsFalse();
        await Assert.That(PartialPuts.Owns(FssHttpCells.EditorsTablePartition)).IsFalse();
    }

    /// <summary>
    /// A store that counts the bytes read from it and written to it.
    /// </summary>
    sealed class CountingStore :
        IFssHttpStore
    {
        public MemoryFssHttpStore Inner { get; } = new();

        public int Read { get; set; }

        public int Written { get; set; }

        public async Task<FssHttpState?> Get(string fileId, Guid partition, Cancel cancel = default)
        {
            var state = await Inner.Get(fileId, partition, cancel);
            Read += state?.Data.Length ?? 0;
            return state;
        }

        public Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FssHttpState value, Cancel cancel = default)
        {
            Written += value.Data.Length;
            return Inner.CompareExchange(fileId, partition, comparand, value, cancel);
        }

        public Task Delete(string fileId, Cancel cancel = default) =>
            Inner.Delete(fileId, cancel);
    }

    /// <summary>
    /// What waits for a file, of all its users, is bounded: a part beyond the limit is refused, and nothing of it kept.
    /// </summary>
    [Test]
    public async Task Add_Bounded()
    {
        var (puts, store, _) = Start();
        var partition = FssHttpCells.ContentPartition;
        await Assert.That(await puts.Add("file", "alice", partition, [Blob(1, 50_000)], 80_000)).IsTrue();
        var stored = await Stored(store);
        await Assert.That(await puts.Add("file", "bob", partition, [Blob(2, 40_000)], 80_000)).IsFalse();
        await Assert.That(await puts.Add("file", "alice", partition, [Blob(3, 40_000)], 80_000)).IsFalse();
        await Assert.That(await Stored(store)).IsEqualTo(stored);
        await Assert.That(await puts.Get("file", "bob", partition)).IsEmpty();
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(1);

        // within it, and without one
        await Assert.That(await puts.Add("file", "bob", partition, [Blob(2, 20_000)], 80_000)).IsTrue();
        await Assert.That(await puts.Add("file", "bob", partition, [Blob(4, 40_000)])).IsTrue();
        await Assert.That(await puts.Get("file", "bob", partition)).Count().IsEqualTo(2);
    }

    /// <summary>
    /// A value of the store that cannot be read is of no use to a last part: nothing waits, and the next write replaces
    /// it, rather than every split save of the file failing on it.
    /// </summary>
    [Test]
    public async Task Unreadable_Replaced()
    {
        var (puts, store, _) = Start();
        var partition = FssHttpCells.ContentPartition;
        await store.CompareExchange("file", PartialPuts.Partition, null, new("1", new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }));
        await Assert.That(await puts.Get("file", "alice", partition)).IsEmpty();

        await puts.Clean("file", "alice", partition, null);
        await Assert.That((await store.Get("file", PartialPuts.Partition))!.Version).IsNotEqualTo("1");
        await Assert.That(await puts.Add("file", "alice", partition, [Blob(1, 1000)])).IsTrue();
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(1);
    }
}
