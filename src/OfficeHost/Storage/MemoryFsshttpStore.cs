namespace OfficeHost;

/// <summary>
/// The states of desktop Office coauthoring in memory, lost on restart.
/// </summary>
public sealed class MemoryFsshttpStore :
    IFsshttpStore
{
    Lock sync = new();
    readonly Dictionary<string, Dictionary<Guid, FsshttpState>> files = [];

    public Task<FsshttpState?> Get(string fileId, Guid partition, Cancel cancel = default)
    {
        lock (sync)
        {
            return Task.FromResult(Current(fileId, partition));
        }
    }

    public Task<bool> CompareExchange(string fileId, Guid partition, string? comparand, FsshttpState value, Cancel cancel = default)
    {
        lock (sync)
        {
            if (Current(fileId, partition)?.Version != comparand)
            {
                return Task.FromResult(false);
            }

            if (!files.TryGetValue(fileId, out var partitions))
            {
                partitions = [];
                files[fileId] = partitions;
            }

            partitions[partition] = value;
            return Task.FromResult(true);
        }
    }

    public Task Delete(string fileId, Cancel cancel = default)
    {
        lock (sync)
        {
            files.Remove(fileId);
            return Task.CompletedTask;
        }
    }

    FsshttpState? Current(string fileId, Guid partition)
    {
        if (files.TryGetValue(fileId, out var partitions) &&
            partitions.TryGetValue(partition, out var state))
        {
            return state;
        }

        return null;
    }
}
