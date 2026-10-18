/// <summary>
/// A local storage where each Reva token is a distinct user, as the CS3 storage maps them.
/// The local storage itself maps every user to root.
/// </summary>
sealed class MultiUserStorage(IStorage inner) :
    IStorage
{
    public (string UserId, string WopiUser) GetUserIdFromCreds(string token, string wopiUser) =>
        (token, $"{wopiUser}!{token}");

    public Task<string> HealthCheck() =>
        inner.HealthCheck();

    public Task<StatInfo> Stat(string endpoint, string filePath, string userId) =>
        inner.Stat(endpoint, filePath, userId);

    public Task<StatInfo> StatX(string endpoint, string fileRef, string userId) =>
        inner.StatX(endpoint, fileRef, userId);

    public Task SetXattr(string endpoint, string filePath, string userId, string key, string value, LockMetadata? lockMetadata) =>
        inner.SetXattr(endpoint, filePath, userId, key, value, lockMetadata);

    public Task RemoveXattr(string endpoint, string filePath, string userId, string key, LockMetadata? lockMetadata) =>
        inner.RemoveXattr(endpoint, filePath, userId, key, lockMetadata);

    public Task SetLock(string endpoint, string filePath, string userId, string appName, string? value) =>
        inner.SetLock(endpoint, filePath, userId, appName, value);

    public Task<StorageLock?> GetLock(string endpoint, string filePath, string userId) =>
        inner.GetLock(endpoint, filePath, userId);

    public Task RefreshLock(string endpoint, string filePath, string userId, string appName, string? value, string? oldValue = null) =>
        inner.RefreshLock(endpoint, filePath, userId, appName, value, oldValue);

    public Task Unlock(string endpoint, string filePath, string userId, string appName, string? value) =>
        inner.Unlock(endpoint, filePath, userId, appName, value);

    public Task<Stream> ReadFile(string endpoint, string filePath, string userId, string? lockId) =>
        inner.ReadFile(endpoint, filePath, userId, lockId);

    public Task WriteFile(string endpoint, string filePath, string userId, Stream content, long? size, LockMetadata? lockMetadata, bool isLock = false, bool noVersion = false) =>
        inner.WriteFile(endpoint, filePath, userId, content, size, lockMetadata, isLock, noVersion);

    public Task RenameFile(string endpoint, string origFilePath, string newFilePath, string userId, LockMetadata? lockMetadata) =>
        inner.RenameFile(endpoint, origFilePath, newFilePath, userId, lockMetadata);

    public Task RemoveFile(string endpoint, string filePath, string userId, bool force = false) =>
        inner.RemoveFile(endpoint, filePath, userId, force);
}
