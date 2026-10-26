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
}
