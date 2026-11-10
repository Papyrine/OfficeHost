/// <summary>
/// The data elements of the parts of a save waiting for its last part (<see cref="PartialPuts"/>): those of a save that
/// never ended do not stay in the store.
/// </summary>
public class PartialPutsTests
{
    static readonly Guid guid = FssHttpB.Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");

    static ObjectDataBlob Blob(uint id, int length) =>
        new(new(guid, id), new(guid, id), new byte[length]);

    static (PartialPuts Puts, MemoryFssHttpStore Store, ManualTime Time) Start()
    {
        var store = new MemoryFssHttpStore();
        var time = new ManualTime();
        return (new(store, time), store, time);
    }

    static async Task<int> Stored(MemoryFssHttpStore store) =>
        (await store.Get("file", PartialPuts.Partition))?.Data.Length ?? 0;

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
        await puts.Clean("file", "carol", partition, last: false);
        await Assert.That(await puts.Get("file", "alice", partition)).Count().IsEqualTo(1);
        await Assert.That(await Stored(store)).IsGreaterThan(60_000);

        // alice's have waited more than an hour
        time.Advance(TimeSpan.FromMinutes(20));
        await Assert.That(await puts.Get("file", "alice", partition)).IsEmpty();
        await Assert.That(await Stored(store)).IsGreaterThan(60_000);
        await puts.Clean("file", "carol", partition, last: false);
        await Assert.That(await Stored(store)).IsLessThan(20_000);
        await Assert.That(await puts.Get("file", "bob", partition)).Count().IsEqualTo(1);

        // bob's last part applied
        await puts.Clean("file", "bob", partition, last: true);
        await Assert.That(await puts.Get("file", "bob", partition)).IsEmpty();
        await Assert.That(await Stored(store)).IsLessThan(100);
    }
}
