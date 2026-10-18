/// <summary>
/// The result of a stat call.
/// </summary>
/// <param name="Inode">A file id invariant to save operations, as produced by <see cref="Common.EncodeInode"/>.</param>
/// <param name="FilePath">The path of the file, usable to reference it in further calls.</param>
/// <param name="MTime">The modification time, in Unix seconds.</param>
public record StatInfo(
    string Inode,
    string FilePath,
    string OwnerId,
    long Size,
    long MTime,
    string Etag,
    IReadOnlyDictionary<string, string> Xattrs);

/// <summary>
/// The app holding a lock and the (encoded) lock id, passed to operations that must honour the lock.
/// </summary>
public record LockMetadata(string AppName, string? LockId);

/// <summary>
/// A Reva-compliant lock, serialized as
/// <c>{"lock_id": "id1234", "type": 2, "user": {...}, "app_name": "your_app", "expiration": {"seconds": 1665446400}}</c>.
/// </summary>
public class StorageLock
{
    [JsonPropertyName("lock_id")]
    public string? LockId { get; set; }

    /// <summary>
    /// LOCK_TYPE_WRITE.
    /// </summary>
    [JsonPropertyName("type")]
    public int Type { get; set; } = 2;

    [JsonPropertyName("app_name")]
    public string AppName { get; set; } = "";

    [JsonPropertyName("user")]
    public JsonObject User { get; set; } = new();

    [JsonPropertyName("expiration")]
    public LockExpiration Expiration { get; set; } = new();
}

public class LockExpiration
{
    [JsonPropertyName("seconds")]
    public long Seconds { get; set; }
}

/// <summary>
/// A failure of the storage layer. See <see cref="IStorage"/> for how the message is interpreted.
/// </summary>
public class StorageException(string message, Exception? inner = null) :
    IOException(message, inner)
{
    public bool IsNotFound => Message.Contains(Common.EnoentMsg);

    public bool IsLockMismatch => Message.Contains(Common.ExclError);

    public bool IsAccessDenied => Message == Common.AccessError;
}
