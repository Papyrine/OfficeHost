/// <summary>
/// The items waiting to be sent, bounded: exchanges up to a count (their bytes are bounded by the
/// <see cref="TapBudget"/>), tombstones up to another, beyond which they are only counted. And the requests in flight,
/// started and not yet ended, for the watermark: no exchange started before it is still to come.
/// </summary>
sealed class TapQueue(int capacity = 1000, int tombstoneCapacity = 10_000)
{
    Lock gate = new();
    List<TapItem> items = [];
    int tombstones;
    long dropped;
    Dictionary<long, DateTime> inFlight = [];

    public int Count
    {
        get
        {
            lock (gate)
            {
                return items.Count;
            }
        }
    }

    public void Started(long sequence, DateTime started)
    {
        lock (gate)
        {
            inFlight[sequence] = started;
        }
    }

    public void Ended(long sequence)
    {
        lock (gate)
        {
            inFlight.Remove(sequence);
        }
    }

    /// <summary>
    /// Adds an exchange, or returns false when the queue is full: it becomes a tombstone.
    /// </summary>
    public bool TryAdd(TapItem item)
    {
        lock (gate)
        {
            if (items.Count - tombstones >= capacity)
            {
                return false;
            }

            items.Add(item);
            return true;
        }
    }

    /// <summary>
    /// Adds a tombstone, or counts it as dropped when there are too many already.
    /// </summary>
    public void AddTombstone(TapItem item)
    {
        lock (gate)
        {
            if (tombstones >= tombstoneCapacity)
            {
                dropped++;
                return;
            }

            items.Add(item);
            tombstones++;
        }
    }

    /// <summary>
    /// The items of the next batch, up to a count and a number of bytes (at least one), with the watermark: the earliest
    /// start of the items left and the requests in flight, else now. The dropped count is taken with them.
    /// </summary>
    public List<TapItem> Take(int maximum, long maximumBytes, out DateTime watermark, out long droppedSince)
    {
        lock (gate)
        {
            var batch = new List<TapItem>();
            long bytes = 0;
            foreach (var item in items)
            {
                if (batch.Count == maximum ||
                    (batch.Count > 0 && bytes + item.Length > maximumBytes))
                {
                    break;
                }

                batch.Add(item);
                bytes += item.Length;
            }

            items.RemoveRange(0, batch.Count);
            tombstones -= batch.Count(_ => _.IsTombstone);
            watermark = items.Select(_ => _.Started)
                .Concat(inFlight.Values)
                .DefaultIfEmpty(TapClock.UtcNow)
                .Min();
            droppedSince = dropped;
            dropped = 0;
            return batch;
        }
    }

    /// <summary>
    /// Puts a batch that was not sent back at the front, as it was.
    /// </summary>
    public void Return(List<TapItem> batch, long droppedSince)
    {
        lock (gate)
        {
            items.InsertRange(0, batch);
            tombstones += batch.Count(_ => _.IsTombstone);
            dropped += droppedSince;
        }
    }

    /// <summary>
    /// Turns every exchange waiting into a tombstone, its copies given back: the shadow cannot be reached.
    /// </summary>
    public void Bury(string reason)
    {
        lock (gate)
        {
            for (var index = 0; index < items.Count; index++)
            {
                if (!items[index].IsTombstone)
                {
                    items[index] = items[index].Bury(reason);
                    tombstones++;
                }
            }
        }
    }
}
