/// <summary>
/// The data elements of partial Put Changes ([MS-FSSHTTPB] 2.2.2.1.4, Partial and Partial Last). Word splits a save of
/// more than about 3 MB of changes: Put Changes of Object Groups alone, each with a Storage Index that maps nothing and
/// no expected one, then a last one with the rest and the Storage Index of the save (docs/fsshttp-notes.md, "SharePoint
/// Server 2016"). The spec says nothing of how a server ties them together, and nothing in a part names its save: the
/// data elements of each part wait here, by file, user and partition, until a last Put Changes of the user applies
/// them with its own, or for an hour from their part. A last part takes the parts its save holds data elements of, and
/// leaves the others: they are of another save of the user, as of Word on a second machine, or of a save that never
/// ended, and go once their hour has passed, with the next Put Changes of the file (<see cref="Clean"/>).
/// <para>
/// Kept in the <see cref="IFssHttpStore"/> as <see cref="SaveAsFiles"/> are, in partitions of their own: the list of
/// the parts in one (<see cref="Partition"/>), and the data elements of each part in one of its own, a slot, so that a
/// part costs the store its own size and not that of the parts before it, and a Put Changes with none to forget only
/// the list. The store cannot forget a partition of a file: the slot of a part that is gone is emptied, and taken
/// again by a later part. The list is written before the slot, so that no two parts take one slot; a part whose slot
/// does not have its version (its write failed) has no data elements.
/// </para>
/// </summary>
class PartialPuts(IFssHttpStore store, TimeProvider time)
{
    /// <summary>
    /// The partition of the list of the parts in the <see cref="IFssHttpStore"/>, which no Office partition uses.
    /// </summary>
    public static Guid Partition { get; } = new("5C1D7A3E-2F84-4B6A-9E0D-8A3B6C4F1E27");

    // what the partitions of the slots have in common, after the number of the slot
    static byte[] slots = PartitionOf(0).ToByteArray()[4..];

    static TimeSpan lifetime = TimeSpan.FromHours(1);

    const int maxAttempts = 8;

    // "OHP" and the version of the format
    static ReadOnlySpan<byte> magic => "OHP2"u8;

    /// <summary>
    /// The partition of the data elements of the part in a slot.
    /// </summary>
    public static Guid PartitionOf(int slot) =>
        new(slot, 0x2F85, 0x4B6A, 0x9E, 0x0D, 0x8A, 0x3B, 0x6C, 0x4F, 0x1E, 0x28);

    /// <summary>
    /// Whether the partition is one of these, the list or a slot: not a cell that a client may query or put.
    /// </summary>
    public static bool Owns(Guid partition) =>
        partition == Partition ||
        partition.ToByteArray().AsSpan(4).SequenceEqual(slots);

    /// <summary>
    /// Adds the data elements of a partial Put Changes of the user to those waiting, as a part with an hour of its own:
    /// a part sent again replaces the one waiting. False, with nothing added, when the parts waiting for the file, of
    /// all its users, would be more than <paramref name="limit"/> bytes: what a client may leave in the store is
    /// bounded.
    /// </summary>
    public async Task<bool> Add(string fileId, string userId, Guid partition, IReadOnlyList<DataElement> elements, long? limit = null, Cancel cancel = default)
    {
        var added = new Dictionary<ExtendedGuid, DataElement>();
        foreach (var element in elements)
        {
            added[element.Id] = element;
        }

        if (added.Count == 0)
        {
            return true;
        }

        var data = Write(added.Values);
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, parts, gone, _) = await Read(fileId, cancel);
            // the same data elements again, or more: the part before is of no more use
            var replaced = parts
                .Where(_ => _.UserId == userId && _.Partition == partition && _.Ids.All(added.ContainsKey))
                .ToList();
            parts.RemoveAll(replaced.Contains);
            if (parts.Sum(_ => _.Size) + data.Length > limit)
            {
                return false;
            }

            var part = new Part(userId, partition, time.GetUtcNow() + lifetime, FreeSlot(parts), NewVersion(), data.Length, added.Keys.ToList());
            parts.Add(part);
            if (!await store.CompareExchange(fileId, Partition, version, new(NewVersion(), Write(parts)), cancel))
            {
                continue;
            }

            await Fill(fileId, part, data, cancel);
            await Empty(fileId, replaced.Concat(gone).Where(_ => _.Slot != part.Slot), cancel);
            return true;
        }

        throw new InvalidOperationException($"The partial Put Changes of {fileId} kept changing");
    }

    /// <summary>
    /// The data elements of the partial Put Changes of the user waiting for a last one, none when there are none: of a
    /// data element in two parts, that of the later.
    /// </summary>
    public async Task<IReadOnlyList<DataElement>> Get(string fileId, string userId, Guid partition, Cancel cancel = default)
    {
        var (_, parts, _, _) = await Read(fileId, cancel);
        var waiting = new Dictionary<ExtendedGuid, DataElement>();
        foreach (var part in parts.Where(_ => _.UserId == userId && _.Partition == partition))
        {
            var stored = await store.Get(fileId, PartitionOf(part.Slot), cancel);
            if (stored?.Version != part.Version)
            {
                continue;
            }

            foreach (var element in Elements(stored.Data))
            {
                waiting[element.Id] = element;
            }
        }

        return waiting.Values.ToList();
    }

    /// <summary>
    /// After a Put Changes of the file applied: forgets the parts of the user that <paramref name="applied"/> holds
    /// data elements of, the cell storage a last part made with them (null for any other Put Changes), and those of any
    /// user that waited more than an hour, whose save never ended. Nothing else would: a file with none to forget
    /// costs a read of the list.
    /// </summary>
    public async Task Clean(string fileId, string userId, Guid partition, CellStorage? applied, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, parts, gone, unreadable) = await Read(fileId, cancel);
            if (applied != null)
            {
                var taken = parts
                    .Where(_ => _.UserId == userId && _.Partition == partition && _.Ids.Any(applied.Elements.ContainsKey))
                    .ToList();
                parts.RemoveAll(taken.Contains);
                gone.AddRange(taken);
            }

            if (gone.Count == 0 &&
                !unreadable)
            {
                return;
            }

            if (await store.CompareExchange(fileId, Partition, version, new(NewVersion(), Write(parts)), cancel))
            {
                await Empty(fileId, gone, cancel);
                return;
            }
        }
    }

    static string NewVersion() => Guid.NewGuid().ToString("N");

    /// <summary>
    /// The first slot that no part of the list has.
    /// </summary>
    static int FreeSlot(List<Part> parts)
    {
        var taken = parts.Select(_ => _.Slot).ToHashSet();
        var slot = 0;
        while (taken.Contains(slot))
        {
            slot++;
        }

        return slot;
    }

    /// <summary>
    /// Writes the data elements of a part of the list to its slot, whatever the slot had: the part before it there is
    /// gone from the list, and may be emptied meanwhile.
    /// </summary>
    async Task Fill(string fileId, Part part, byte[] data, Cancel cancel)
    {
        var slot = PartitionOf(part.Slot);
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var version = await store.GetVersion(fileId, slot, cancel);
            if (await store.CompareExchange(fileId, slot, version, new(part.Version, data), cancel))
            {
                return;
            }
        }

        throw new InvalidOperationException($"The data elements of a partial Put Changes of {fileId} could not be stored");
    }

    /// <summary>
    /// Empties the slots of parts gone from the list, unless a later part has taken one already.
    /// </summary>
    async Task Empty(string fileId, IEnumerable<Part> gone, Cancel cancel)
    {
        foreach (var part in gone)
        {
            await store.CompareExchange(fileId, PartitionOf(part.Slot), part.Version, new(NewVersion(), Array.Empty<byte>()), cancel);
        }
    }

    /// <summary>
    /// The list: the parts that wait, those whose hour has passed, and whether the store has a value that cannot be
    /// read, which is of no use to a last part and is replaced by the next write.
    /// </summary>
    async Task<(string? Version, List<Part> Parts, List<Part> Gone, bool Unreadable)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, Partition, cancel);
        if (stored == null)
        {
            return (null, [], [], false);
        }

        List<Part> all;
        try
        {
            all = Parse(stored.Data);
        }
        catch (InvalidDataException)
        {
            return (stored.Version, [], [], true);
        }

        var now = time.GetUtcNow();
        return (stored.Version, all.Where(_ => _.Expires > now).ToList(), all.Where(_ => _.Expires <= now).ToList(), false);
    }

    /// <summary>
    /// The magic and format version, the number of parts, then for each its user as a String Item, its partition, its
    /// expiry in UTC ticks, its slot, the version of its slot as a String Item, the size of its data elements, and their
    /// ids as an Extended GUID Array ([MS-FSSHTTPB] 2.2.1).
    /// </summary>
    static byte[] Write(List<Part> parts)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        writer.WriteBytes(magic);
        writer.WriteCompactUInt64((ulong) parts.Count);
        foreach (var part in parts)
        {
            writer.WriteStringItem(part.UserId);
            writer.WriteGuid(part.Partition);
            writer.WriteCompactUInt64((ulong) part.Expires.UtcTicks);
            writer.WriteCompactUInt64((ulong) part.Slot);
            writer.WriteStringItem(part.Version);
            writer.WriteCompactUInt64((ulong) part.Size);
            writer.WriteExtendedGuidArray(part.Ids);
        }

        return output.WrittenSpan.ToArray();
    }

    static List<Part> Parse(ReadOnlyMemory<byte> data)
    {
        try
        {
            var reader = new FssHttpBReader(data);
            if (!reader.ReadBytes(4).Span.SequenceEqual(magic))
            {
                throw new InvalidDataException("Not the partial Put Changes of a file");
            }

            var count = reader.ReadCompactUInt64();
            var parts = new List<Part>();
            for (ulong index = 0; index < count; index++)
            {
                var userId = reader.ReadStringItem();
                var partition = reader.ReadGuid();
                var expires = new DateTimeOffset((long) reader.ReadCompactUInt64(), TimeSpan.Zero);
                var slot = checked((int) reader.ReadCompactUInt64());
                var version = reader.ReadStringItem();
                var size = checked((long) reader.ReadCompactUInt64());
                parts.Add(new(userId, partition, expires, slot, version, size, reader.ReadExtendedGuidArray()));
            }

            reader.EnsureEnd("the partial Put Changes of a file");
            return parts;
        }
        catch (Exception exception) when (exception is FssHttpBException or ArgumentOutOfRangeException or OverflowException)
        {
            throw new InvalidDataException($"The partial Put Changes of a file are corrupt: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// The data elements of a part, as a Data Element Package ([MS-FSSHTTPB] 2.2.1.12).
    /// </summary>
    static byte[] Write(IEnumerable<DataElement> elements)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        new DataElementPackage(elements.ToList()).Write(ref writer);
        return output.WrittenSpan.ToArray();
    }

    static IReadOnlyList<DataElement> Elements(ReadOnlyMemory<byte> data)
    {
        try
        {
            var reader = new FssHttpBReader(data);
            var elements = DataElementPackage.Read(ref reader).Elements;
            reader.EnsureEnd("the data elements of a partial Put Changes");
            return elements;
        }
        catch (FssHttpBException exception)
        {
            throw new InvalidDataException($"The data elements of a partial Put Changes are corrupt: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// A partial Put Changes that waits: whose, until when, and where its data elements are, the slot with the version
    /// it has while it holds them, their size and their ids.
    /// </summary>
    record Part(string UserId, Guid Partition, DateTimeOffset Expires, int Slot, string Version, long Size, IReadOnlyList<ExtendedGuid> Ids);
}
