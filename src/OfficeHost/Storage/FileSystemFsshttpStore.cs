namespace OfficeHost;

/// <summary>
/// The states of desktop Office coauthoring in a folder, for a single process: a folder per file, named by the SHA-256
/// of its id, with a file per partition, which holds the version then the state. A state is written to a temporary
/// file then moved over the previous one, so a crash leaves either.
/// </summary>
public sealed class FileSystemFsshttpStore :
    IFsshttpStore
{
    SemaphoreSlim sync = new(1, 1);

    /// <param name="root">The folder of the states, created if missing.</param>
    public FileSystemFsshttpStore(string root)
    {
        Root = Path.GetFullPath(root);
        Directory.CreateDirectory(Root);
    }

    /// <summary>
    /// The full path of the folder of the states.
    /// </summary>
    public string Root { get; }

    string FolderOf(string fileId) =>
        Path.Combine(Root, Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fileId)))[..32]);

    string PathOf(string fileId, Guid partition) => Path.Combine(FolderOf(fileId), $"{partition:N}.state");

    public async Task<FsshttpState?> Get(string fileId, Guid partition, Cancel cancel = default)
    {
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

    public async Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FsshttpState value, Cancel cancel = default)
    {
        await sync.WaitAsync(cancel);
        try
        {
            var path = PathOf(fileId, partition);
            var current = await Read(path, cancel);
            if (current?.Version != comparand)
            {
                return false;
            }

            Directory.CreateDirectory(FolderOf(fileId));
            var version = Encoding.UTF8.GetBytes(value.Version);
            var temporary = $"{path}.{Guid.NewGuid():N}.tmp";
            await using (var stream = File.Create(temporary))
            {
                var length = new byte[4];
                BinaryPrimitives.WriteInt32LittleEndian(length, version.Length);
                await stream.WriteAsync(length, cancel);
                await stream.WriteAsync(version, cancel);
                await stream.WriteAsync(value.Data, cancel);
            }

            File.Move(temporary, path, true);
            return true;
        }
        finally
        {
            sync.Release();
        }
    }

    public async Task Delete(string fileId, Cancel cancel = default)
    {
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

    static async Task<FsshttpState?> Read(string path, Cancel cancel)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        var bytes = await File.ReadAllBytesAsync(path, cancel);
        if (bytes.Length < 4)
        {
            throw new InvalidDataException($"The state {path} is truncated");
        }

        var length = BinaryPrimitives.ReadInt32LittleEndian(bytes);
        if (length < 0 || length > bytes.Length - 4)
        {
            throw new InvalidDataException($"The state {path} is truncated");
        }

        var version = Encoding.UTF8.GetString(bytes, 4, length);
        return new(version, bytes.AsMemory(4 + length));
    }
}
