/// <summary>
/// A storage failing the operations a test picks, to reach the error handling of the server.
/// </summary>
sealed class FaultyStorage(IStorage inner) :
    ForwardingStorage(inner)
{
    /// <summary>
    /// Given an operation name (the IStorage method) and a path, the exception to throw, or null to go on.
    /// </summary>
    public Func<string, string, Exception?> Fault { get; set; } = (_, _) => null;

    /// <summary>
    /// Changes what stat returns, for instance to give an EOS etag.
    /// </summary>
    public Func<StatInfo, StatInfo> MapStat { get; set; } = _ => _;

    void Check(string operation, string path)
    {
        var exception = Fault(operation, path);
        if (exception != null)
        {
            throw exception;
        }
    }

    public override async Task<StatInfo> Stat(string endpoint, string filePath, string userId)
    {
        Check(nameof(Stat), filePath);
        return MapStat(await base.Stat(endpoint, filePath, userId));
    }

    public override async Task<StatInfo> StatX(string endpoint, string fileRef, string userId)
    {
        Check(nameof(StatX), fileRef);
        return MapStat(await base.StatX(endpoint, fileRef, userId));
    }

    public override Task SetXattr(string endpoint, string filePath, string userId, string key, string value, LockMetadata? lockMetadata)
    {
        Check(nameof(SetXattr), filePath);
        return base.SetXattr(endpoint, filePath, userId, key, value, lockMetadata);
    }

    public override Task SetLock(string endpoint, string filePath, string userId, string appName, string? value)
    {
        Check(nameof(SetLock), filePath);
        return base.SetLock(endpoint, filePath, userId, appName, value);
    }

    public override Task<StorageLock?> GetLock(string endpoint, string filePath, string userId)
    {
        Check(nameof(GetLock), filePath);
        return base.GetLock(endpoint, filePath, userId);
    }

    public override Task RefreshLock(string endpoint, string filePath, string userId, string appName, string? value, string? oldValue = null)
    {
        Check(nameof(RefreshLock), filePath);
        return base.RefreshLock(endpoint, filePath, userId, appName, value, oldValue);
    }

    public override Task Unlock(string endpoint, string filePath, string userId, string appName, string? value)
    {
        Check(nameof(Unlock), filePath);
        return base.Unlock(endpoint, filePath, userId, appName, value);
    }

    public override Task<Stream> ReadFile(string endpoint, string filePath, string userId, string? lockId)
    {
        Check(nameof(ReadFile), filePath);
        return base.ReadFile(endpoint, filePath, userId, lockId);
    }

    public override Task WriteFile(string endpoint, string filePath, string userId, Stream content, long? size, LockMetadata? lockMetadata, bool isLock = false, bool noVersion = false)
    {
        Check(nameof(WriteFile), filePath);
        return base.WriteFile(endpoint, filePath, userId, content, size, lockMetadata, isLock, noVersion);
    }

    public override Task RenameFile(string endpoint, string origFilePath, string newFilePath, string userId, LockMetadata? lockMetadata)
    {
        Check(nameof(RenameFile), origFilePath);
        return base.RenameFile(endpoint, origFilePath, newFilePath, userId, lockMetadata);
    }

    public override Task RemoveFile(string endpoint, string filePath, string userId, bool force = false)
    {
        Check(nameof(RemoveFile), filePath);
        return base.RemoveFile(endpoint, filePath, userId, force);
    }
}
