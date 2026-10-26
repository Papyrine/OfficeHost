/// <summary>
/// The cell storages of the files, kept in the <see cref="IFssHttpStore"/>. The content partition is kept in step with
/// the file in the <see cref="IOfficeStorage"/>: when the file changed since the cell was written (through WOPI, or by
/// the application), its bytes are written as the next revision of the cell, which keeps the chunks that did not change,
/// so that Office downloads only the others.
/// </summary>
class FssHttpCells(IFssHttpStore store, IOfficeStorage storage, FssHttpCellCache cache)
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
        var key = (fileId, partition);
        if (cache.TryGet(key, out var cached) &&
            await store.GetVersion(fileId, partition, cancel) == cached.Version)
        {
            return cached;
        }

        var stored = await store.Get(fileId, partition, cancel);
        if (stored == null)
        {
            cache.Remove(key);
            return new(null, CellState.New());
        }

        var cell = new StoredCell(stored.Version, CellState.Read(stored.Data));
        cache.Set(key, cell);
        return cell;
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

            if (cell.State is { Storage: { } pending, FileVersion: null } &&
                cell.State.BaseFileVersion == file.Version)
            {
                // a Put Changes stored the cell, but did not get to write the file, which has not changed since
                return await WriteFile(fileId, cell, FileCell.Read(pending).Root.ToArray(), cancel);
            }

            // otherwise a cell the file left behind, or that was not written to the file before the file changed, is
            // written again from the file

            var bytes = await ReadFile(fileId, cancel);
            var ids = cell.State.Ids();
            CellUpdate update;
            try
            {
                update = FileCell.Write(bytes, cell.State.Storage, ids, ChunkingOptions.Default);
            }
            catch (FssHttpDException)
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

    /// <summary>
    /// Writes the bytes of a content cell stored without a file version to the file, then stores the cell with the
    /// version of the file. When another write stored the cell meanwhile, that one is kept.
    /// </summary>
    public async Task<StoredCell> WriteFile(string fileId, StoredCell cell, byte[] bytes, Cancel cancel = default)
    {
        // the file is written whatever the request becomes: a write left halfway would leave the cell pending
        var file = await storage.Write(fileId, new MemoryStream(bytes), Cancel.None);
        var state = cell.State with
        {
            FileVersion = file.Version,
            BaseFileVersion = null
        };
        return await TryUpdate(fileId, ContentPartition, cell.Version, state, cancel) ??
               cell with
               {
                   State = state
               };
    }

    async Task<byte[]> ReadFile(string fileId, Cancel cancel)
    {
        await using var stream = await storage.OpenRead(fileId, cancel);
        using var memory = new MemoryStream();
        await stream.CopyToAsync(memory, cancel);
        return memory.ToArray();
    }

    /// <summary>
    /// Replaces the cell of a partition with one that holds nothing yet and allocates from new ids: the content
    /// partition is then written again from the file.
    /// </summary>
    public async Task Reset(string fileId, Guid partition, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var cell = await Get(fileId, partition, cancel);
            if (await TryUpdate(fileId, partition, cell.Version, CellState.New(), cancel) != null)
            {
                return;
            }
        }

        throw new InvalidOperationException($"The cell of the file {fileId} changed {maxAttempts} times while it was reset");
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

        var cell = new StoredCell(version, state);
        cache.Set((fileId, partition), cell);
        return cell;
    }
}
