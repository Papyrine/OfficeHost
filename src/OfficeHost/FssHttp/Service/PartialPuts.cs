/// <summary>
/// The data elements of partial Put Changes ([MS-FSSHTTPB] 2.2.2.1.4, Partial and Partial Last). Word splits a save of
/// more than about 3 MB of changes: Put Changes of Object Groups alone, each with a Storage Index that maps nothing and
/// no expected one, then a last one with the rest and the Storage Index of the save (docs/fsshttp-notes.md, "SharePoint
/// Server 2016"). The spec says nothing of how a server ties them together, and nothing in a part names its save: the
/// data elements of each part wait here, by file, user and partition, until a last Put Changes of the user applies
/// them with its own, or for an hour from their part. A last part takes those its save holds, and leaves the others:
/// they are of another save of the user, as of Word on a second machine, or of a save that never ended, and go once
/// their hour has passed, with the next Put Changes of the file (<see cref="Clean"/>). Kept in the
/// <see cref="IFssHttpStore"/> as <see cref="SaveAsFiles"/> are, in a partition of their own.
/// </summary>
class PartialPuts(IFssHttpStore store, TimeProvider time)
{
    /// <summary>
    /// The partition of the data elements in the <see cref="IFssHttpStore"/>, which no Office partition uses.
    /// </summary>
    public static Guid Partition { get; } = new("5C1D7A3E-2F84-4B6A-9E0D-8A3B6C4F1E27");

    static TimeSpan lifetime = TimeSpan.FromHours(1);

    const int maxAttempts = 8;

    // "OHP" and the version of the format
    static ReadOnlySpan<byte> magic => "OHP1"u8;

    /// <summary>
    /// Adds the data elements of a partial Put Changes of the user to those waiting, as a part with an hour of its own:
    /// one of the same id replaces the one waiting. False, with nothing added, when the parts waiting for the file, of
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

        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, parts, _) = await Read(fileId, cancel);
            Remove(parts, userId, partition, added.ContainsKey);
            parts.Add(new(userId, partition, time.GetUtcNow() + lifetime, added.Values.ToList()));
            var data = Write(parts);
            if (data.Length > limit)
            {
                return false;
            }

            if (await store.CompareExchange(fileId, Partition, version, new(Guid.NewGuid().ToString("N"), data), cancel))
            {
                return true;
            }
        }

        throw new InvalidOperationException($"The partial Put Changes of {fileId} kept changing");
    }

    /// <summary>
    /// The data elements of the partial Put Changes of the user waiting for a last one, none when there are none.
    /// </summary>
    public async Task<IReadOnlyList<DataElement>> Get(string fileId, string userId, Guid partition, Cancel cancel = default)
    {
        var (_, parts, _) = await Read(fileId, cancel);
        return parts
            .Where(_ => _.UserId == userId && _.Partition == partition)
            .SelectMany(_ => _.Elements)
            .ToList();
    }

    /// <summary>
    /// After a Put Changes of the file applied: forgets the data elements of the user that <paramref name="applied"/>
    /// holds, the cell storage a last part made with them (null for any other Put Changes), and those of any user that
    /// waited more than an hour, whose save never ended. Nothing else would: a file with none waiting costs a read.
    /// </summary>
    public async Task Clean(string fileId, string userId, Guid partition, CellStorage? applied, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, parts, stale) = await Read(fileId, cancel);
            var removed = applied != null &&
                          Remove(parts, userId, partition, applied.Elements.ContainsKey);
            if (!removed && !stale ||
                await store.CompareExchange(fileId, Partition, version, new(Guid.NewGuid().ToString("N"), Write(parts)), cancel))
            {
                return;
            }
        }
    }

    /// <summary>
    /// Takes the data elements that match out of the parts of the user for the partition, and the parts left with none
    /// out of the list. Whether it took any.
    /// </summary>
    static bool Remove(List<Part> parts, string userId, Guid partition, Func<ExtendedGuid, bool> match)
    {
        var removed = false;
        for (var index = parts.Count - 1; index >= 0; index--)
        {
            var part = parts[index];
            if (part.UserId != userId ||
                part.Partition != partition)
            {
                continue;
            }

            var kept = part.Elements.Where(_ => !match(_.Id)).ToList();
            if (kept.Count == part.Elements.Count)
            {
                continue;
            }

            removed = true;
            if (kept.Count == 0)
            {
                parts.RemoveAt(index);
                continue;
            }

            parts[index] = part with
            {
                Elements = kept
            };
        }

        return removed;
    }

    /// <summary>
    /// The parts that have not expired, and whether the store has some to drop: expired, or a value that cannot be read,
    /// which is of no use to a last part and is replaced by the next write.
    /// </summary>
    async Task<(string? Version, List<Part> Parts, bool Stale)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, Partition, cancel);
        if (stored == null)
        {
            return (null, [], false);
        }

        List<Part> all;
        try
        {
            all = Parse(stored.Data);
        }
        catch (InvalidDataException)
        {
            return (stored.Version, [], true);
        }

        var now = time.GetUtcNow();
        var parts = all.Where(_ => _.Expires > now).ToList();
        return (stored.Version, parts, parts.Count < all.Count);
    }

    /// <summary>
    /// The magic and format version, the number of parts, then for each its user as a String Item, its partition, its
    /// expiry in UTC ticks, and its data elements as a Data Element Package ([MS-FSSHTTPB] 2.2.1).
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
            new DataElementPackage(part.Elements).Write(ref writer);
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
                parts.Add(new(userId, partition, expires, DataElementPackage.Read(ref reader).Elements));
            }

            reader.EnsureEnd("the partial Put Changes of a file");
            return parts;
        }
        catch (Exception exception) when (exception is FssHttpBException or ArgumentOutOfRangeException)
        {
            throw new InvalidDataException($"The partial Put Changes of a file are corrupt: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// The data elements of one partial Put Changes, until a last part takes them or they expire.
    /// </summary>
    record Part(string UserId, Guid Partition, DateTimeOffset Expires, IReadOnlyList<DataElement> Elements);
}
