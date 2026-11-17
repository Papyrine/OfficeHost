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

        // the sequences received, in the order received, to know an exchange of a batch sent again
        HashSet<long> sequences = [];
        Queue<(long Sequence, DateTimeOffset Received)> received = new();

        public bool Has(long sequence) => sequences.Contains(sequence);

        public void Received(long sequence, DateTimeOffset now)
        {
            if (sequences.Add(sequence))
            {
                received.Enqueue((sequence, now));
            }
        }

        /// <summary>
        /// Forgets the sequences received before the time: a batch is not sent again after so long.
        /// </summary>
        public void Forget(DateTimeOffset before)
        {
            while (received.TryPeek(out var oldest) &&
                   oldest.Received < before)
            {
                received.Dequeue();
                sequences.Remove(oldest.Sequence);
            }
        }
    }

    Lock gate = new();
    List<ShadowItem> pending = [];
    long pendingBytes;
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
    /// Whether the inbox takes no exchange more: so many wait, or so many bytes of them, that a batch is refused before
    /// it is read.
    /// </summary>
    public bool Full
    {
        get
        {
            var settings = options.Value;
            lock (gate)
            {
                return pending.Count >= settings.QueueCapacity ||
                       pendingBytes >= MaxBytes(settings);
            }
        }
    }

    static long MaxBytes(OfficeHostShadowOptions settings) =>
        settings.QueueMegabytes * 1024L * 1024;

    /// <summary>
    /// The bytes an item holds: the bodies of its exchange, a tombstone none.
    /// </summary>
    static long Bytes(ShadowItem item)
    {
        if (item.Exchange is not { } exchange)
        {
            return 0;
        }

        return exchange.Request.Body.Length + exchange.Response.Body.Length;
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
    /// Adds a batch, or refuses it when the inbox is full, of <see cref="OfficeHostShadowOptions.QueueCapacity"/>
    /// exchanges or <see cref="OfficeHostShadowOptions.QueueMegabytes"/>: the module backs off, and drops with
    /// tombstones. A batch larger than the megabytes is taken by an empty inbox. Of a batch sent again, as when the
    /// answer to it was lost, only the exchanges not received are added, so that none is replayed twice.
    /// </summary>
    public bool Add(ShadowBatch batch)
    {
        var now = time.GetUtcNow();
        var settings = options.Value;
        lock (gate)
        {
            instances.TryGetValue(batch.Instance, out var instance);
            var items = batch.Items
                .Where(_ => instance == null || !instance.Has(_.Sequence))
                .ToList();
            var bytes = items.Sum(Bytes);
            if (items.Count > 0 &&
                (pending.Count + items.Count > settings.QueueCapacity ||
                 (pending.Count > 0 && pendingBytes + bytes > MaxBytes(settings))))
            {
                refused += items.Count;
                return false;
            }

            if (instance == null)
            {
                instance = new(batch.Server);
                instances[batch.Instance] = instance;
            }

            foreach (var item in items)
            {
                instance.Received(item.Sequence, now);
            }

            if (batch.Watermark > instance.Watermark)
            {
                instance.Watermark = batch.Watermark;
            }

            instance.LastSeen = now;
            instance.Final |= batch.Final;
            instance.Dropped += batch.Dropped;
            instance.Lost = false;
            received += items.Count;
            pending.AddRange(items.Select(_ => _ with { Received = now }));
            pendingBytes += bytes;
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
            // an exchange that waited long enough is released with those that started before it, which may have come
            // after it: released later, they would be replayed after it
            var waited = DateTimeOffset.MinValue;
            foreach (var item in pending)
            {
                if (now - item.Received >= settings.MaximumDelay &&
                    item.Started > waited)
                {
                    waited = item.Started;
                }
            }

            bool Ready(ShadowItem item) =>
                item.Started <= waited ||
                watermarks.All(_ => _ >= item.Started);

            var ready = pending
                .Where(Ready)
                .OrderBy(_ => _.Started)
                .ThenBy(_ => _.Sequence)
                .ToList();
            // in one pass: the inbox holds up to QueueCapacity
            pending.RemoveAll(Ready);
            pendingBytes -= ready.Sum(Bytes);
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
                instance.Forget(now - settings.IdleTimeout);
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
