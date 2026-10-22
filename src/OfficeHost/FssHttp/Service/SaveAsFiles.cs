/// <summary>
/// The files desktop sites made by Save As. A site stands for the file of its token; Word saves as a file of another name
/// in the same folder, at the URL of that name in the site, and goes on editing it there (docs/fsshttp-notes.md). The
/// site reaches those files only, not the others of the folder: they are kept with the state of the file of the token in
/// the <see cref="IFssHttpStore"/>, by a hash of the token, until the token expires.
/// </summary>
class SaveAsFiles(IFssHttpStore store, TimeProvider time)
{
    /// <summary>
    /// The partition of the files in the <see cref="IFssHttpStore"/>, which no Office partition uses.
    /// </summary>
    public static Guid Partition { get; } = new("0E6B2F7C-58A1-4D3E-9B64-3F8C1A2D7E15");

    const int maxAttempts = 8;

    static JsonSerializerOptions json = new(JsonSerializerDefaults.Web);

    /// <summary>
    /// The ids of the files the site made, or took over, by Save As.
    /// </summary>
    public async Task<IReadOnlyList<string>> Get(FssHttpSite site, Cancel cancel = default)
    {
        var (_, entries) = await Read(site.Access.FileId, cancel);
        var token = Hash(site.Token);
        return entries
            .Where(_ => _.Token == token)
            .Select(_ => _.FileId)
            .ToList();
    }

    public async Task Add(FssHttpSite site, string fileId, Cancel cancel = default)
    {
        var token = Hash(site.Token);
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, entries) = await Read(site.Access.FileId, cancel);
            if (entries.Any(_ => _.Token == token && _.FileId == fileId))
            {
                return;
            }

            entries.Add(new(token, fileId, site.Expires));
            var data = JsonSerializer.SerializeToUtf8Bytes(entries, json);
            if (await store.CompareExchange(site.Access.FileId, Partition, version, new(Guid.NewGuid().ToString("N"), data), cancel))
            {
                return;
            }
        }

        throw new InvalidOperationException($"The Save As files of {site.Access.FileId} kept changing");
    }

    async Task<(string? Version, List<Entry> Entries)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, Partition, cancel);
        if (stored == null)
        {
            return (null, []);
        }

        var now = time.GetUtcNow();
        var entries = JsonSerializer.Deserialize<List<Entry>>(stored.Data.Span, json) ?? [];
        return (stored.Version, entries.Where(_ => _.Expires > now).ToList());
    }

    static string Hash(string token) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(token)));

    record Entry(string Token, string FileId, DateTimeOffset Expires);
}
