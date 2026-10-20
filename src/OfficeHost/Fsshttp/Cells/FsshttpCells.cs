/// <summary>
/// A cell state as stored, with the version of the <see cref="FsshttpState"/> it came from, null when there was none.
/// </summary>
record StoredCell(string? Version, CellState State);

/// <summary>
/// The cell storages of the files, kept in the <see cref="IFsshttpStore"/>. The content partition is kept in step with
/// the file in the <see cref="IOfficeStorage"/>: when the file changed since the cell was written (through WOPI, or by
/// the application), its bytes are written as the next revision of the cell, which keeps the chunks that did not change,
/// so that Office downloads only the others.
/// </summary>
class FsshttpCells(IFsshttpStore store, IOfficeStorage storage)
{
    /// <summary>
    /// The partitions of [MS-FSSHTTP] 2.3.3.1 (PartitionID): the content of the file, its metadata, and the editors
    /// table.
    /// </summary>
    public static Guid ContentPartition { get; } = Guid.Empty;

    public static Guid MetadataPartition { get; } = new("383ADC0B-E66E-4438-95E6-E39EF9720122");

    public static Guid EditorsTablePartition { get; } = new("7808F4DD-2385-49D6-B7CE-37ACA5E43602");

    const int maxAttempts = 8;

    public async Task<StoredCell> Get(string fileId, Guid partition, Cancel cancel = default)
    {
        var stored = await store.Get(fileId, partition, cancel);
        if (stored == null)
        {
            return new(null, CellState.New());
        }

        return new(stored.Version, CellState.Read(stored.Data));
    }

    /// <summary>
    /// The content partition, written again from the file when the file has another version. Null when there is no
    /// file. A state whose cell is not a file stream is replaced by one written from nothing: Office then downloads the
    /// whole file.
    /// </summary>
    public async Task<StoredCell?> GetContent(string fileId, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var file = await storage.GetFile(fileId, cancel);
            if (file == null)
            {
                return null;
            }

            var cell = await Get(fileId, ContentPartition, cancel);
            if (cell.State.Storage != null &&
                cell.State.FileVersion == file.Version)
            {
                return cell;
            }

            var bytes = await ReadFile(fileId, cancel);
            var ids = cell.State.Ids();
            CellUpdate update;
            try
            {
                update = FileCell.Write(bytes, cell.State.Storage, ids, ChunkingOptions.Default);
            }
            catch (FsshttpdException)
            {
                update = FileCell.Write(bytes, null, ids, ChunkingOptions.Default);
            }

            var updated = await TryUpdate(fileId, ContentPartition, cell.Version, new(update.Storage, ids.Guid, ids.Next, file.Version), cancel);
            if (updated != null)
            {
                return updated;
            }
        }

        throw new InvalidOperationException($"The cell of the file {fileId} changed {maxAttempts} times while it was written");
    }

    async Task<byte[]> ReadFile(string fileId, Cancel cancel)
    {
        await using var stream = await storage.OpenRead(fileId, cancel);
        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory, cancel);
        return memory.ToArray();
    }

    /// <summary>
    /// Stores the state when the stored one still has the version <paramref name="comparand"/>, with a new version.
    /// Null when another write came first.
    /// </summary>
    public async Task<StoredCell?> TryUpdate(string fileId, Guid partition, string? comparand, CellState state, Cancel cancel = default)
    {
        var version = Guid.NewGuid().ToString("N");
        if (!await store.CompareExchange(fileId, partition, comparand, new(version, state.ToArray()), cancel))
        {
            return null;
        }

        return new(version, state);
    }
}
