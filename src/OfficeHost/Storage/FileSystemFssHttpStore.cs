namespace OfficeHost;

/// <summary>
/// The states of desktop Office coauthoring in a folder, for a single process: a folder per file, named by the SHA-256
/// of its id, with a file per partition, which holds the version then the state. A state is written to a temporary
/// file then moved over the previous one, so a crash leaves either. The operations on a file take turns, those on
/// different files do not.
/// </summary>
public sealed class FileSystemFssHttpStore :
    IFssHttpStore
{
    // a lock per file folder, among a fixed number so that they need no cleanup
    SemaphoreSlim[] syncs = Enumerable.Range(0, 64).Select(_ => new SemaphoreSlim(1, 1)).ToArray();

    /// <param name="root">The folder of the states, created if missing.</param>
    public FileSystemFssHttpStore(string root)
    {
        Root = Path.GetFullPath(root);
        Directory.CreateDirectory(Root);
    }

    /// <summary>
    /// The full path of the folder of the states.
    /// </summary>
    public string Root { get; }

    static string FolderName(string fileId) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fileId)))[..32];

    string FolderOf(string fileId) => Path.Combine(Root, FolderName(fileId));

    string PathOf(string fileId, Guid partition) => Path.Combine(FolderOf(fileId), $"{partition:N}.state");

    SemaphoreSlim SyncOf(string fileId) =>
        syncs[Convert.ToInt32(FolderName(fileId)[..4], 16) % syncs.Length];

    public async Task<FssHttpState?> Get(string fileId, Guid partition, Cancel cancel = default)
    {
        var sync = SyncOf(fileId);
        await sync.WaitAsync(cancel);
        try
        {
            return await Read(PathOf(fileId, partition), cancel);
        }
        finally
        {
            sync.Release();
        }
    }

    public async Task<string?> GetVersion(string fileId, Guid partition, Cancel cancel = default)
    {
        var sync = SyncOf(fileId);
        await sync.WaitAsync(cancel);
        try
        {
            return await ReadVersion(PathOf(fileId, partition), cancel);
        }
        finally
        {
            sync.Release();
        }
    }

    public async Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FssHttpState value, Cancel cancel = default)
    {
        var sync = SyncOf(fileId);
        await sync.WaitAsync(cancel);
        try
        {
            var path = PathOf(fileId, partition);
            if (await ReadVersion(path, cancel) != comparand)
            {
                return false;
            }

            Directory.CreateDirectory(FolderOf(fileId));
            var version = Encoding.UTF8.GetBytes(value.Version);
            var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
            try
            {
                await using (var stream = File.Create(temporary))
                {
                    var length = new byte[4];
                    BinaryPrimitives.WriteInt32LittleEndian(length, version.Length);
                    await stream.WriteAsync(length, cancel);
                    await stream.WriteAsync(version, cancel);
                    await stream.WriteAsync(value.Data, cancel);
                }

                File.Move(temporary, path, true);
            }
            catch
            {
                File.Delete(temporary);
                throw;
            }

            return true;
        }
        finally
        {
            sync.Release();
        }
    }

    public async Task Delete(string fileId, Cancel cancel = default)
    {
        var sync = SyncOf(fileId);
        await sync.WaitAsync(cancel);
        try
        {
            var folder = FolderOf(fileId);
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
        }
        finally
        {
            sync.Release();
        }
    }

    static async Task<FssHttpState?> Read(string path, Cancel cancel)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        var bytes = await File.ReadAllBytesAsync(path, cancel);
        var length = VersionLength(bytes, bytes.Length, path);
        var version = Encoding.UTF8.GetString(bytes, 4, length);
        return new(version, bytes.AsMemory(4 + length));
    }

    /// <summary>
    /// The version of the state only, without reading the state.
    /// </summary>
    static async Task<string?> ReadVersion(string path, Cancel cancel)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
        var header = new byte[4];
        if (stream.Length < header.Length)
        {
            throw new InvalidDataException($"The state {path} is truncated");
        }

        await stream.ReadExactlyAsync(header, cancel);
        var length = VersionLength(header, stream.Length, path);
        var version = new byte[length];
        await stream.ReadExactlyAsync(version, cancel);
        return Encoding.UTF8.GetString(version);
    }

    static int VersionLength(byte[] header, long fileLength, string path)
    {
        if (fileLength < 4 ||
            header.Length < 4)
        {
            throw new InvalidDataException($"The state {path} is truncated");
        }

        var length = BinaryPrimitives.ReadInt32LittleEndian(header);
        if (length < 0 ||
            length > fileLength - 4)
        {
            throw new InvalidDataException($"The state {path} is truncated");
        }

        return length;
    }
}
