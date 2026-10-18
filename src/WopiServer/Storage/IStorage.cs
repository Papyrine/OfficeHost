/// <summary>
/// A storage layer for the WOPI server. Failures are reported with a <see cref="StorageException"/> whose
/// message contains <see cref="Common.EnoentMsg"/>, <see cref="Common.ExclError"/> or
/// <see cref="Common.AccessError"/> when applicable: the WOPI logic classifies errors by those messages.
/// </summary>
public interface IStorage
{
    /// <summary>
    /// Probes the storage and returns a status message.
    /// </summary>
    Task<string> HealthCheck();

    /// <summary>
    /// Maps a Reva token and a wopiuser to the credentials used to access the storage, and to a wopiuser
    /// in the form `username!userid_as_returned_by_stat`.
    /// </summary>
    (string UserId, string WopiUser) GetUserIdFromCreds(string token, string wopiUser);

    /// <summary>
    /// Stats a file on behalf of the given user. Fails with `Is a directory` for folders.
    /// </summary>
    Task<StatInfo> Stat(string endpoint, string filePath, string userId);

    /// <summary>
    /// Gets the extended stat info, where <paramref name="fileRef"/> may be a path or a storage-specific file id.
    /// </summary>
    Task<StatInfo> StatX(string endpoint, string fileRef, string userId);

    Task SetXattr(string endpoint, string filePath, string userId, string key, string value, LockMetadata? lockMetadata);

    Task RemoveXattr(string endpoint, string filePath, string userId, string key, LockMetadata? lockMetadata);

    /// <summary>
    /// Atomically sets the lock, failing with <see cref="Common.ExclError"/> if a valid lock is already present.
    /// </summary>
    Task SetLock(string endpoint, string filePath, string userId, string appName, string? value);

    /// <summary>
    /// Returns the current lock, or null when the file is not locked or the lock has expired.
    /// </summary>
    Task<StorageLock?> GetLock(string endpoint, string filePath, string userId);

    /// <summary>
    /// Refreshes the lock, replacing <paramref name="oldValue"/> (the current lock when null) with <paramref name="value"/>.
    /// </summary>
    Task RefreshLock(string endpoint, string filePath, string userId, string appName, string? value, string? oldValue = null);

    Task Unlock(string endpoint, string filePath, string userId, string appName, string? value);

    /// <summary>
    /// Opens the file for reading. The caller owns the returned stream.
    /// </summary>
    Task<Stream> ReadFile(string endpoint, string filePath, string userId, string? lockId);

    /// <summary>
    /// Writes the whole content, replacing any pre-existing one. <paramref name="size"/> is the expected length of
    /// the content, when known. With <paramref name="isLock"/> the file is created in O_CREAT|O_EXCL mode, failing
    /// with <see cref="Common.ExclError"/> when it exists. With <paramref name="noVersion"/> no new version is created
    /// where the storage supports versioning.
    /// </summary>
    Task WriteFile(string endpoint, string filePath, string userId, Stream content, long? size, LockMetadata? lockMetadata, bool isLock = false, bool noVersion = false);

    Task RenameFile(string endpoint, string origFilePath, string newFilePath, string userId, LockMetadata? lockMetadata);

    /// <summary>
    /// Removes a file. With <paramref name="force"/> the file bypasses any recycle bin, where applicable.
    /// </summary>
    Task RemoveFile(string endpoint, string filePath, string userId, bool force = false);
}
