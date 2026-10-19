/// <summary>
/// Keeps the WOPI locks. The built-in <see cref="MemoryLockStore"/> suits a single instance of the application; a
/// store shared by all instances is needed otherwise.
/// </summary>
public interface IWopiLockStore
{
    /// <summary>
    /// The lock of the file, or null when it is not locked or the lock expired.
    /// </summary>
    Task<WopiLock?> Get(string fileId, Cancel cancel = default);

    /// <summary>
    /// Atomically replaces the lock of the file with <paramref name="value"/> (null to unlock) when its current lock id
    /// is <paramref name="comparand"/> (null for an unlocked file), and returns the lock it had before, as
    /// <see cref="Interlocked.CompareExchange{T}(ref T, T, T)"/> does. An expired lock counts as no lock.
    /// </summary>
    Task<WopiLock?> CompareExchange(string fileId, string? comparand, WopiLock? value, Cancel cancel = default);
}

/// <param name="Id">The lock id the WOPI client chose, compared as is.</param>
public record WopiLock(string Id, DateTimeOffset Expires);

/// <summary>
/// The locks in memory, lost on restart.
/// </summary>
public sealed class MemoryLockStore(TimeProvider time) :
    IWopiLockStore
{
    readonly Lock sync = new();
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
