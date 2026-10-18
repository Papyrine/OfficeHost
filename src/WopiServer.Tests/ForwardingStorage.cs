/// <summary>
/// Forwards every call to another storage: tests override what they simulate.
/// </summary>
class ForwardingStorage(IStorage inner) :
    IStorage
{
    public virtual (string UserId, string WopiUser) GetUserIdFromCreds(string token, string wopiUser) =>
        inner.GetUserIdFromCreds(token, wopiUser);

    public virtual Task<string> HealthCheck() =>
        inner.HealthCheck();

    public virtual Task<StatInfo> Stat(string endpoint, string filePath, string userId) =>
        inner.Stat(endpoint, filePath, userId);

    public virtual Task<StatInfo> StatX(string endpoint, string fileRef, string userId) =>
        inner.StatX(endpoint, fileRef, userId);

    public virtual Task SetXattr(string endpoint, string filePath, string userId, string key, string value, LockMetadata? lockMetadata) =>
        inner.SetXattr(endpoint, filePath, userId, key, value, lockMetadata);

    public virtual Task RemoveXattr(string endpoint, string filePath, string userId, string key, LockMetadata? lockMetadata) =>
        inner.RemoveXattr(endpoint, filePath, userId, key, lockMetadata);

    public virtual Task SetLock(string endpoint, string filePath, string userId, string appName, string? value) =>
        inner.SetLock(endpoint, filePath, userId, appName, value);

    public virtual Task<StorageLock?> GetLock(string endpoint, string filePath, string userId) =>
        inner.GetLock(endpoint, filePath, userId);

    public virtual Task RefreshLock(string endpoint, string filePath, string userId, string appName, string? value, string? oldValue = null) =>
        inner.RefreshLock(endpoint, filePath, userId, appName, value, oldValue);

    public virtual Task Unlock(string endpoint, string filePath, string userId, string appName, string? value) =>
        inner.Unlock(endpoint, filePath, userId, appName, value);

    public virtual Task<Stream> ReadFile(string endpoint, string filePath, string userId, string? lockId) =>
        inner.ReadFile(endpoint, filePath, userId, lockId);

    public virtual Task WriteFile(string endpoint, string filePath, string userId, Stream content, long? size, LockMetadata? lockMetadata, bool isLock = false, bool noVersion = false) =>
        inner.WriteFile(endpoint, filePath, userId, content, size, lockMetadata, isLock, noVersion);

    public virtual Task RenameFile(string endpoint, string origFilePath, string newFilePath, string userId, LockMetadata? lockMetadata) =>
        inner.RenameFile(endpoint, origFilePath, newFilePath, userId, lockMetadata);

    public virtual Task RemoveFile(string endpoint, string filePath, string userId, bool force = false) =>
        inner.RemoveFile(endpoint, filePath, userId, force);
}
