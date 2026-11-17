/// <summary>
/// The order <see cref="ShadowInbox"/> releases the exchanges in, across the instances of the module.
/// </summary>
public class ShadowInboxTests
{
    static readonly Guid first = Guid.NewGuid();
    static readonly Guid second = Guid.NewGuid();

    static (ShadowInbox Inbox, ManualTime Time) Start()
    {
        var time = new ManualTime();
        var options = Microsoft.Extensions.Options.Options.Create(
            new OfficeHostShadowOptions
            {
                MaximumDelay = TimeSpan.FromSeconds(15)
            });
        return (new(options, time), time);
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
}
