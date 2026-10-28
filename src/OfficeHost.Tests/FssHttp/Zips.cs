/// <summary>
/// Zip files for the chunking tests, made with <see cref="ZipArchive"/> on a seekable stream, so that the sizes are in
/// the local file headers, with a fixed time, so that the same items make the same bytes.
/// </summary>
static class Zips
{
    static readonly DateTimeOffset time = new(2026, 10, 20, 12, 0, 0, TimeSpan.Zero);

    public static byte[] Create(params (string Name, byte[] Data)[] items) => Create(CompressionLevel.Optimal, items);

    public static byte[] Create(CompressionLevel level, params (string Name, byte[] Data)[] items)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            foreach (var (name, data) in items)
            {
                var entry = archive.CreateEntry(name, level);
                entry.LastWriteTime = time;
                using var entryStream = entry.Open();
                entryStream.Write(data);
            }
        }

        return stream.ToArray();
    }

    public static Dictionary<string, string> Read(byte[] zip)
    {
        using var archive = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(_ => _.FullName, _ =>
        {
            using var reader = new StreamReader(_.Open());
            return reader.ReadToEnd();
        });
    }

    public static byte[] Random(int length, int seed)
    {
        var bytes = new byte[length];
        new Random(seed).NextBytes(bytes);
        return bytes;
    }

    public static byte[] Text(string text) => Encoding.UTF8.GetBytes(text);
}
