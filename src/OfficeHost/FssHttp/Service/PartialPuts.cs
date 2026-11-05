/// <summary>
/// The data elements of partial Put Changes ([MS-FSSHTTPB] 2.2.2.1.4, Partial and Partial Last). Word splits a save of
/// more than about 3 MB of changes: Put Changes of Object Groups alone, each with a Storage Index that maps nothing and
/// no expected one, then a last one with the rest and the Storage Index of the save (docs/fsshttp-notes.md, "SharePoint
/// Server 2016"). The spec says nothing of how a server ties them together: the data elements wait here, by file, user
/// and partition, until the last Put Changes applies them with its own, or for an hour. Kept in the
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
    /// Adds the data elements of a partial Put Changes of the user to those waiting: one of the same id replaces the
    /// one there.
    /// </summary>
    public async Task Add(string fileId, string userId, Guid partition, IReadOnlyList<DataElement> elements, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, entries) = await Read(fileId, cancel);
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
        var (_, entries) = await Read(fileId, cancel);
        return entries.FirstOrDefault(_ => _.UserId == userId && _.Partition == partition)?.Elements ?? [];
    }

    /// <summary>
    /// Forgets the data elements of the user, once the last Put Changes applied them.
    /// </summary>
    public async Task Remove(string fileId, string userId, Guid partition, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, entries) = await Read(fileId, cancel);
            if (entries.RemoveAll(_ => _.UserId == userId && _.Partition == partition) == 0 ||
                await store.CompareExchange(fileId, Partition, version, new(Guid.NewGuid().ToString("N"), Write(entries)), cancel))
            {
                return;
            }
        }
    }

    async Task<(string? Version, List<Entry> Entries)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, Partition, cancel);
        if (stored == null)
        {
            return (null, []);
        }

        var now = time.GetUtcNow();
        return (stored.Version, Parse(stored.Data).Where(_ => _.Expires > now).ToList());
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
