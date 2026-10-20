namespace OfficeHost;

/// <summary>
/// Keeps the WOPI locks. The built-in <see cref="MemoryLockStore"/> suits a single instance of the application; a
/// store shared by all instances is needed otherwise.
/// </summary>
public interface IOfficeLockStore
{
    /// <summary>
    /// The lock of the file, or null when it is not locked or the lock expired.
    /// </summary>
    Task<OfficeLock?> Get(string fileId, Cancel cancel = default);

    /// <summary>
    /// Atomically replaces the lock of the file with <paramref name="value"/> (null to unlock) when its current lock id
    /// is <paramref name="comparand"/> (null for an unlocked file), and returns the lock it had before, as
    /// <see cref="Interlocked.CompareExchange{T}(ref T, T, T)"/> does. An expired lock counts as no lock.
    /// </summary>
    Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, Cancel cancel = default);
}

/// <param name="Id">The lock id the WOPI client chose, compared as is.</param>
public record OfficeLock(string Id, DateTimeOffset Expires);