/// <summary>
/// Common entities used by all storage interfaces, including the functions to store and retrieve
/// Reva-compatible locks.
/// </summary>
public static class Common
{
    // Standard Linux error messages, as also used by the CS3 client. Storage errors are classified by these
    // messages, so a storage interface must use them verbatim.

    /// <summary>
    /// The file is missing.
    /// </summary>
    public const string EnoentMsg = "No such file or directory";

    /// <summary>
    /// Attempt to overwrite a file or xattr in O_EXCL mode, or a lock operation whose preconditions failed.
    /// </summary>
    public const string ExclError = "Lock mismatch";

    /// <summary>
    /// Attempt an operation without the required access rights.
    /// </summary>
    public const string AccessError = "Operation not permitted";

    /// <summary>
    /// Name of the xattr storing the Reva lock.
    /// </summary>
    public const string LockKey = "iop.lock";

    /// <summary>
    /// The prefix used for the lock-id payload to be WebDAV compatible, see
    /// https://github.com/cs3org/wopiserver/pull/51#issuecomment-1038798545. The UUID is fully random and
    /// hard coded to identify WOPI locks.
    /// </summary>
    public const string WebDavLockPrefix = "opaquelocktoken:797356a8-0500-4ceb-a8a0-c94c8cde7eba";

    /// <summary>
    /// A base64-encoded lock compatible with the Reva implementation of the CS3 Lock API, cf.
    /// https://github.com/cs3org/cs3apis/blob/main/cs3/storage/provider/v1beta1/resources.proto
    /// </summary>
    public static string GenRevaLock(string appName, string? value, long expiration)
    {
        var revaLock = new StorageLock
        {
            LockId = value,
            AppName = appName,
            Expiration = new()
            {
                Seconds = expiration
            }
        };
        return UrlSafeBase64(JsonSerializer.SerializeToUtf8Bytes(revaLock, Json.Options));
    }

    /// <summary>
    /// Restores the JSON payload of a base64-encoded Reva lock.
    /// </summary>
    public static StorageLock RetrieveRevaLock(string rawLock)
    {
        try
        {
            return JsonSerializer.Deserialize<StorageLock>(Base64Url.DecodeFromChars(rawLock.TrimEnd('=')), Json.Options) ??
                   throw new StorageException("Unable to parse existing lock: null lock");
        }
        catch (Exception exception) when (exception is FormatException or JsonException or DecoderFallbackException)
        {
            throw new StorageException($"Unable to parse existing lock: {exception.Message}");
        }
    }

    /// <summary>
    /// Encodes an endpoint and an inode to be used as a safe WOPISrc. The endpoint is assumed to be URL safe and
    /// without `!` characters, so the separator allows the inverse operation (as the web frontend does).
    /// </summary>
    public static string EncodeInode(string endpoint, string inode) =>
        $"{endpoint}!{UrlSafeBase64(Encoding.UTF8.GetBytes(inode))}";

    /// <summary>
    /// Decodes an inode obtained from <see cref="EncodeInode"/>.
    /// </summary>
    public static (string Endpoint, string Inode) DecodeInode(string inode)
    {
        var parts = inode.Split('!');
        if (parts.Length != 2)
        {
            throw new FormatException($"Invalid inode {inode}");
        }

        return (parts[0], Encoding.UTF8.GetString(Base64Url.DecodeFromChars(parts[1].TrimEnd('='))));
    }

    /// <summary>
    /// Python <c>urlsafe_b64encode</c>, which keeps the padding.
    /// </summary>
    static string UrlSafeBase64(byte[] bytes) =>
        Convert.ToBase64String(bytes).Replace('+', '-').Replace('/', '_');
}
