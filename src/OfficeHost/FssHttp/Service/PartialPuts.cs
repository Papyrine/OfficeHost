/// <summary>
/// The data elements of partial Put Changes ([MS-FSSHTTPB] 2.2.2.1.4, Partial and Partial Last). Word splits a save of
/// more than about 3 MB of changes: Put Changes of Object Groups alone, each with a Storage Index that maps nothing and
/// no expected one, then a last one with the rest and the Storage Index of the save (docs/fsshttp-notes.md, "SharePoint
/// Server 2016"). The spec says nothing of how a server ties them together: the data elements wait here, by file, user
/// and partition, until the last Put Changes applies them with its own, or for an hour: those of a save that never
/// ended go with the next Put Changes of the file (<see cref="Clean"/>). Kept in the <see cref="IFssHttpStore"/> as
/// <see cref="SaveAsFiles"/> are, in a partition of their own.
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
    /// Adds the data elements of a partial Put Changes of the user to those waiting: one of the same id replaces the
    /// one there.
    /// </summary>
    public async Task Add(string fileId, string userId, Guid partition, IReadOnlyList<DataElement> elements, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, entries, _) = await Read(fileId, cancel);
            var index = entries.FindIndex(_ => _.UserId == userId && _.Partition == partition);
            var waiting = new Dictionary<ExtendedGuid, DataElement>();
            if (index >= 0)
            {
                waiting = entries[index].Elements.ToDictionary(_ => _.Id);
                entries.RemoveAt(index);
            }

            foreach (var element in elements)
            {
                waiting[element.Id] = element;
            }

            entries.Add(new(userId, partition, time.GetUtcNow() + lifetime, waiting.Values.ToList()));
            if (await store.CompareExchange(fileId, Partition, version, new(Guid.NewGuid().ToString("N"), Write(entries)), cancel))
            {
                return;
            }
        }

        throw new InvalidOperationException($"The partial Put Changes of {fileId} kept changing");
    }

    /// <summary>
    /// The data elements of the partial Put Changes of the user waiting for the last one, none when there are none.
    /// </summary>
    public async Task<IReadOnlyList<DataElement>> Get(string fileId, string userId, Guid partition, Cancel cancel = default)
    {
        var (_, entries, _) = await Read(fileId, cancel);
        return entries.FirstOrDefault(_ => _.UserId == userId && _.Partition == partition)?.Elements ?? [];
    }

    /// <summary>
    /// After a Put Changes of the file applied: forgets the data elements of the user when it was the last part of a
    /// save, which applied them, and those of any user that waited more than an hour, whose save never ended. Nothing
    /// else would: a file with none waiting costs a read.
    /// </summary>
    public async Task Clean(string fileId, string userId, Guid partition, bool last, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, entries, expired) = await Read(fileId, cancel);
            var removed = last ? entries.RemoveAll(_ => _.UserId == userId && _.Partition == partition) : 0;
            if (removed == 0 && !expired ||
                await store.CompareExchange(fileId, Partition, version, new(Guid.NewGuid().ToString("N"), Write(entries)), cancel))
            {
                return;
            }
        }
    }

    /// <summary>
    /// The entries that have not expired, and whether the store has some that have.
    /// </summary>
    async Task<(string? Version, List<Entry> Entries, bool Expired)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, Partition, cancel);
        if (stored == null)
        {
            return (null, [], false);
        }

        var now = time.GetUtcNow();
        var all = Parse(stored.Data);
        var entries = all.Where(_ => _.Expires > now).ToList();
        return (stored.Version, entries, entries.Count < all.Count);
    }

    /// <summary>
    /// The magic and format version, the number of entries, then for each its user as a String Item, its partition, its
    /// expiry in UTC ticks, and its data elements as a Data Element Package ([MS-FSSHTTPB] 2.2.1).
    /// </summary>
    static byte[] Write(List<Entry> entries)
    {
        var output = new ArrayBufferWriter<byte>();
        var writer = new FssHttpBWriter(output);
        writer.WriteBytes(magic);
        writer.WriteCompactUInt64((ulong) entries.Count);
        foreach (var entry in entries)
        {
            writer.WriteStringItem(entry.UserId);
            writer.WriteGuid(entry.Partition);
            writer.WriteCompactUInt64((ulong) entry.Expires.UtcTicks);
            new DataElementPackage(entry.Elements).Write(ref writer);
        }

        return output.WrittenSpan.ToArray();
    }

    static List<Entry> Parse(ReadOnlyMemory<byte> data)
    {
        try
        {
            var reader = new FssHttpBReader(data);
            if (!reader.ReadBytes(4).Span.SequenceEqual(magic))
            {
                throw new InvalidDataException("Not the partial Put Changes of a file");
            }

            var count = reader.ReadCompactUInt64();
            var entries = new List<Entry>();
            for (ulong index = 0; index < count; index++)
            {
                var userId = reader.ReadStringItem();
                var partition = reader.ReadGuid();
                var expires = new DateTimeOffset((long) reader.ReadCompactUInt64(), TimeSpan.Zero);
                entries.Add(new(userId, partition, expires, DataElementPackage.Read(ref reader).Elements));
            }

            reader.EnsureEnd("the partial Put Changes of a file");
            return entries;
        }
        catch (FssHttpBException exception)
        {
            throw new InvalidDataException($"The partial Put Changes of a file are corrupt: {exception.Message}", exception);
        }
    }

    record Entry(string UserId, Guid Partition, DateTimeOffset Expires, IReadOnlyList<DataElement> Elements);
}
