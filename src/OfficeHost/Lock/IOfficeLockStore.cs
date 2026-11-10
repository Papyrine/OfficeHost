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

    /// <summary>
    /// <see cref="CompareExchange(string, string?, OfficeLock?, Cancel)"/>, with the user whose request locks, refreshes
    /// or unlocks, so that a store can keep who holds a lock: the built-in stores keep the user as the
    /// <see cref="OfficeLock.Holder"/> of the lock they write, or, with no user, the holder the value names. A store
    /// that does not override this one forgets the user.
    /// </summary>
    /// <param name="user">The user whose request this is, or null for an exchange of the application's own.</param>
    Task<OfficeLock?> CompareExchange(string fileId, string? comparand, OfficeLock? value, OfficeUser? user, Cancel cancel = default) =>
        CompareExchange(fileId, comparand, value, cancel);
}

/// <param name="Id">The lock id the WOPI client chose, compared as is.</param>
public record OfficeLock(string Id, DateTimeOffset Expires)
{
    /// <summary>
    /// The user whose request last took or refreshed the lock, when the store keeps it (the built-in ones do), for an
    /// application to show who is editing. The lock of a desktop coauthoring session is shared by its editors: the
    /// holder is the one whose request last changed it.
    /// </summary>
    public OfficeUser? Holder { get; init; }
}
