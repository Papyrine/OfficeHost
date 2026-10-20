namespace OfficeHost;

/// <summary>
/// The locks in memory, lost on restart.
/// </summary>
public sealed class MemoryLockStore(TimeProvider time) :
    IOfficeLockStore
{
    Lock sync = new();
    readonly Dictionary<string, OfficeLock> locks = [];
    int operations;

    public Task<OfficeLock?> Get(string fileId, Cancel cancel = default)
    {
        lock (sync)
        {
            return Task.FromResult(Current(fileId));
        }
    }

    public Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, Cancel cancel = default)
    {
        lock (sync)
        {
            RemoveExpired();
            var current = Current(fileId);
            if (current?.Id == comparand)
            {
                if (value == null)
                {
                    locks.Remove(fileId);
                }
                else
                {
                    locks[fileId] = value;
                }
            }

            return Task.FromResult(current);
        }
    }

    OfficeLock? Current(string fileId)
    {
        if (locks.TryGetValue(fileId, out var current) &&
            current.Expires > time.GetUtcNow())
        {
            return current;
        }

        return null;
    }

    /// <summary>
    /// Now and then, forgets the locks of the files that were left locked.
    /// </summary>
    void RemoveExpired()
    {
        if (++operations % 256 != 0)
        {
            return;
        }

        var now = time.GetUtcNow();
        foreach (var (fileId, _) in locks.Where(_ => _.Value.Expires <= now).ToList())
        {
            locks.Remove(fileId);
        }
    }
}