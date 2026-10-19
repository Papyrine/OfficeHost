namespace WopiServer;

/// <summary>
/// The locks in memory, lost on restart.
/// </summary>
public sealed class MemoryLockStore(TimeProvider time) :
    IWopiLockStore
{
    Lock sync = new();
    readonly Dictionary<string, WopiLock> locks = [];
    int operations;

    public Task<WopiLock?> Get(string fileId, Cancel cancel = default)
    {
        lock (sync)
        {
            return Task.FromResult(Current(fileId));
        }
    }

    public Task<WopiLock?> CompareExchange(string fileId, string? comparand, WopiLock? value, Cancel cancel = default)
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

    WopiLock? Current(string fileId)
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