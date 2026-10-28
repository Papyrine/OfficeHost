/// <summary>
/// An exchange or a tombstone of a batch, as SharePoint started it: <see cref="ShadowProtocol"/>.
/// </summary>
record ShadowItem(Guid Instance, string Server, long Sequence, DateTimeOffset Started, Exchange? Exchange, ShadowDocument? Tombstone = null, string? Reason = null)
{
    public DateTimeOffset Received { get; init; }
}

/// <summary>
/// A batch of an instance of the module in SharePoint: <see cref="ShadowProtocol"/>.
/// </summary>
record ShadowBatch(Guid Instance, string Server, DateTimeOffset Watermark, long Dropped, bool Final, List<ShadowItem> Items);

/// <summary>
/// A loss the shadow cannot tie to a document: an instance that stopped sending without its last batch (its app domain
/// stopped abruptly), or exchanges it dropped without a tombstone. Every document it sent is then untrusted.
/// </summary>
record ShadowLoss(Guid Instance, string Server, string Reason);

/// <summary>
/// The exchanges received and not yet replayed, released in the order SharePoint started them, across its servers: an
/// exchange once every active instance's watermark has passed its start (no instance can still send one that started
/// before it), or once it has waited <see cref="OfficeHostShadowOptions.MaximumDelay"/> for a lagging one.
/// </summary>
sealed class ShadowInbox(IOptions<OfficeHostShadowOptions> options, TimeProvider time)
{
    sealed class Instance(string server)
    {
        public string Server { get; } = server;
        public DateTimeOffset Watermark { get; set; }
        public DateTimeOffset LastSeen { get; set; }
        public bool Final { get; set; }
        public long Dropped { get; set; }
        public bool Lost { get; set; }
    }

    Lock gate = new();
    List<ShadowItem> pending = [];
    Dictionary<Guid, Instance> instances = [];
    SemaphoreSlim signal = new(0);
    long received;
    long refused;

    public int Count
    {
        get
        {
            lock (gate)
            {
                return pending.Count;
            }
        }
    }

    /// <summary>
    /// The exchanges received and refused (the inbox full) since the start.
    /// </summary>
    public (long Received, long Refused) Totals
    {
        get
        {
            lock (gate)
            {
                return (received, refused);
            }
        }
    }

    /// <summary>
    /// Adds a batch, or refuses it when the inbox is full: the module backs off, and drops with tombstones.
    /// </summary>
    public bool Add(ShadowBatch batch)
    {
        var now = time.GetUtcNow();
        lock (gate)
        {
            if (batch.Items.Count > 0 &&
                pending.Count + batch.Items.Count > options.Value.QueueCapacity)
            {
                refused += batch.Items.Count;
                return false;
            }

            if (!instances.TryGetValue(batch.Instance, out var instance))
            {
                instance = new(batch.Server);
                instances[batch.Instance] = instance;
            }

            if (batch.Watermark > instance.Watermark)
            {
                instance.Watermark = batch.Watermark;
            }

            instance.LastSeen = now;
            instance.Final |= batch.Final;
            instance.Dropped += batch.Dropped;
            instance.Lost = false;
            received += batch.Items.Count;
            pending.AddRange(batch.Items.Select(_ => _ with { Received = now }));
        }

        signal.Release();
        return true;
    }

    public Task Wait(TimeSpan timeout, Cancel cancel) => signal.WaitAsync(timeout, cancel);

    /// <summary>
    /// The exchanges that can be replayed, in the order SharePoint started them.
    /// </summary>
    public List<ShadowItem> Release()
    {
        var now = time.GetUtcNow();
        var settings = options.Value;
        lock (gate)
        {
            var watermarks = instances.Values
                .Where(_ => !_.Final && now - _.LastSeen < settings.InstanceTimeout)
                .Select(_ => _.Watermark)
                .ToList();
            bool Ready(ShadowItem item) =>
                now - item.Received >= settings.MaximumDelay ||
                watermarks.All(_ => _ >= item.Started);

            var ready = pending
                .Where(Ready)
                .OrderBy(_ => _.Started)
                .ThenBy(_ => _.Sequence)
                .ToList();
            // in one pass: the inbox holds up to QueueCapacity
            pending.RemoveAll(Ready);
            return ready;
        }
    }

    /// <summary>
    /// The losses since the last call, each once; instances gone for long are forgotten.
    /// </summary>
    public List<ShadowLoss> Losses()
    {
        var now = time.GetUtcNow();
        var settings = options.Value;
        var losses = new List<ShadowLoss>();
        lock (gate)
        {
            foreach (var (id, instance) in instances.ToList())
            {
                if (instance.Dropped > 0)
                {
                    losses.Add(new(id, instance.Server, $"{instance.Dropped} exchanges dropped without a tombstone"));
                    instance.Dropped = 0;
                }

                var silent = now - instance.LastSeen;
                if (instance is
                    {
                        Final: false,
                        Lost: false
                    } &&
                    silent >= settings.InstanceTimeout)
                {
                    losses.Add(new(id, instance.Server, "stopped sending without its last batch"));
                    instance.Lost = true;
                }

                if (silent >= settings.IdleTimeout)
                {
                    instances.Remove(id);
                }
            }
        }

        return losses;
    }
}
