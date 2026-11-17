/// <summary>
/// The order <see cref="ShadowInbox"/> releases the exchanges in, across the instances of the module, and what it holds.
/// </summary>
public class ShadowInboxTests
{
    static readonly Guid first = Guid.NewGuid();
    static readonly Guid second = Guid.NewGuid();

    static (ShadowInbox Inbox, ManualTime Time) Start(int queueMegabytes = 1024)
    {
        var time = new ManualTime();
        var options = Microsoft.Extensions.Options.Options.Create(
            new OfficeHostShadowOptions
            {
                MaximumDelay = TimeSpan.FromSeconds(15),
                QueueMegabytes = queueMegabytes
            });
        return (new(options, time), time);
    }

    static ShadowItem Exchange(long sequence, DateTimeOffset started, int bytes) =>
        new(first, "first", sequence, started, new(1, new("POST / HTTP/1.1", [], new byte[bytes]), new("HTTP/1.1 200 OK", [], [])));

    /// <summary>
    /// The inbox holds up to QueueMegabytes of exchanges, beyond which it is full and refuses a batch, as beyond
    /// QueueCapacity; a batch larger than it is taken by an empty inbox; replaying frees what it held.
    /// </summary>
    [Test]
    public async Task Add_QueueMegabytes()
    {
        var (inbox, time) = Start(queueMegabytes: 1);
        var start = time.Now;
        const int large = 600 * 1024;
        await Assert.That(inbox.Add(Batch(first, start, Exchange(1, start, large)))).IsTrue();
        await Assert.That(inbox.Full).IsFalse();
        await Assert.That(inbox.Add(Batch(first, start, Exchange(2, start, large)))).IsFalse();
        await Assert.That(inbox.Add(Batch(first, start, Exchange(3, start, 424 * 1024)))).IsTrue();
        await Assert.That(inbox.Full).IsTrue();

        await Assert.That(inbox.Add(Batch(first, start.AddSeconds(1)))).IsTrue();
        await Assert.That(inbox.Release().Count).IsEqualTo(2);
        await Assert.That(inbox.Full).IsFalse();
        await Assert.That(inbox.Add(Batch(first, start.AddSeconds(1), Exchange(4, start.AddSeconds(1), 2 * 1024 * 1024)))).IsTrue();
        await Assert.That(inbox.Full).IsTrue();
    }

    static ShadowItem Item(Guid instance, long sequence, DateTimeOffset started) =>
        new(instance, instance == first ? "first" : "second", sequence, started, null);

    static ShadowBatch Batch(Guid instance, DateTimeOffset watermark, params ShadowItem[] items) =>
        new(instance, instance == first ? "first" : "second", watermark, 0, false, [..items]);

    /// <summary>
    /// An exchange that waited for an instance whose watermark lags is released with the exchanges that started before
    /// it, though they came after it, in their order: released later, they would be replayed after it.
    /// </summary>
    [Test]
    public async Task Release_WaitedWithEarlier()
    {
        var (inbox, time) = Start();
        var start = time.Now;
        // the second instance has an exchange in flight since start: its watermark stays there
        inbox.Add(Batch(second, start));
        var later = Item(first, 1, start.AddSeconds(1));
        inbox.Add(Batch(first, start.AddSeconds(1), later));
        time.Advance(TimeSpan.FromSeconds(1));
        var earlier = Item(second, 1, start.AddSeconds(0.5));
        inbox.Add(Batch(second, start, earlier));
        await Assert.That(inbox.Release()).IsEmpty();

        // the later one has waited, the earlier one not yet
        time.Advance(TimeSpan.FromSeconds(14));
        var released = inbox.Release();
        await Assert.That(released.Select(_ => _.Started).SequenceEqual([earlier.Started, later.Started])).IsTrue();
        await Assert.That(inbox.Count).IsEqualTo(0);
    }

    /// <summary>
    /// A batch sent again, as when the answer to it was lost, adds only the exchanges not received, whether those were
    /// replayed since or still wait: none is replayed twice.
    /// </summary>
    [Test]
    public async Task Add_SentAgain()
    {
        var (inbox, time) = Start();
        var start = time.Now;
        var one = Item(first, 1, start);
        var two = Item(first, 2, start.AddSeconds(1));
        await Assert.That(inbox.Add(Batch(first, start.AddSeconds(2), one))).IsTrue();
        await Assert.That(inbox.Add(Batch(first, start.AddSeconds(2), one, two))).IsTrue();
        await Assert.That(inbox.Count).IsEqualTo(2);
        await Assert.That(inbox.Release().Select(_ => _.Sequence).SequenceEqual([1L, 2L])).IsTrue();

        await Assert.That(inbox.Add(Batch(first, start.AddSeconds(2), one, two))).IsTrue();
        await Assert.That(inbox.Count).IsEqualTo(0);
        await Assert.That(inbox.Totals.Received).IsEqualTo(2);
    }
}
