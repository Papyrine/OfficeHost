/// <summary>
/// The cells <see cref="FssHttpCells"/> read or wrote last, by file and partition, as they were parsed: one is read
/// again only when the version of its state changed. Desktop Office polls the content of a file, whose cell is the
/// size of the file, every few seconds per coauthor.
/// </summary>
class FssHttpCellCache
{
    /// <summary>
    /// The most cells kept: those of the files being coauthored.
    /// </summary>
    const int maxCells = 32;

    ConcurrentDictionary<(string FileId, Guid Partition), StoredCell> cells = new();

    public bool TryGet((string FileId, Guid Partition) key, [NotNullWhen(true)] out StoredCell? cell) =>
        cells.TryGetValue(key, out cell);

    public void Set((string FileId, Guid Partition) key, StoredCell cell)
    {
        // a crude bound, enough for a cache of the few files being coauthored
        if (cells.Count >= maxCells &&
            !cells.ContainsKey(key))
        {
            cells.Clear();
        }

        cells[key] = cell;
    }

    public void Remove((string FileId, Guid Partition) key) =>
        cells.TryRemove(key, out _);

    // a gate per file among a fixed number, so that they need no cleanup
    SemaphoreSlim[] rewrites = Enumerable.Range(0, 16).Select(_ => new SemaphoreSlim(1, 1)).ToArray();

    /// <summary>
    /// The gate of the writes of the content cell of a file from the file: one request of the process at a time reads
    /// the whole file and chunks it, and those polling the file meanwhile find the cell written.
    /// </summary>
    public SemaphoreSlim RewriteGate(string fileId) =>
        rewrites[(fileId.GetHashCode() & int.MaxValue) % rewrites.Length];
}
