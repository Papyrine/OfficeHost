/// <summary>
/// Local storage interface, meant for development and test purposes only.
/// The Python implementation keeps the WOPI metadata (xattrs and locks) in Linux user extended attributes.
/// To work on any OS, this one keeps them in a JSON document per file, under a hidden <c>.wopimeta</c>
/// folder at the storage root, keyed by the file's path.
/// </summary>
public sealed class LocalStorage(WopiConfig config, ILogger log, TimeProvider time) :
    IStorage
{
    const string metaFolderName = ".wopimeta";

    static readonly StringComparison pathComparison =
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    readonly string homePath = ValidateHomePath(config, log);
    readonly ConcurrentDictionary<string, object> gates = new();

    static string ValidateHomePath(WopiConfig config, ILogger log)
    {
        var homePath = config.GetRequired("local", "storagehomepath");
        if (File.Exists(homePath))
        {
            throw new StorageException($"Could not stat storagehomepath folder {homePath}: Not a directory");
        }

        if (!Directory.Exists(homePath))
        {
            throw new StorageException($"Could not stat storagehomepath folder {homePath}: {Common.EnoentMsg}");
        }

        // all right but inform the user
        log.LogWarning("Use this local storage interface for test/development purposes only, not for production");
        return Path.TrimEndingDirectorySeparator(Path.GetFullPath(homePath));
    }

    string MetaPath => Path.Combine(homePath, metaFolderName);

    long Now => time.GetUtcNow().ToUnixTimeSeconds();

    int ChunkSize => Math.Max(config.GetInt("io", "chunksize"), 1);

    /// <summary>
    /// Maps the given storage path into the local file system, below the storagehomepath.
    /// </summary>
    string GetFilePath(string filePath)
    {
        var full = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.Join(homePath, filePath)));
        if (!IsBelow(full, homePath) ||
            IsBelow(full, MetaPath))
        {
            log.LogWarning("Attempt to access a path outside the storage {filepath}", filePath);
            throw new StorageException(Common.AccessError);
        }

        return full;
    }

    static bool IsBelow(string path, string folder) =>
        string.Equals(path, folder, pathComparison) ||
        path.StartsWith(Path.TrimEndingDirectorySeparator(folder) + Path.DirectorySeparatorChar, pathComparison);

    public async Task<string> HealthCheck()
    {
        try
        {
            await Stat("", "/", "");
            // to please CodeQL but never reached
            return "Warning";
        }
        catch (StorageException exception)
        {
            if (exception.Message == "Is a directory")
            {
                // that's expected, yet we return warning as this is a test/dev storage interface
                log.LogDebug("Executed health check against storage root");
                return "Warning";
            }

            // any other error is a failure
            log.LogError("Health check failed against storage root {error}", exception.Message);
            return exception.Message;
        }
    }

    /// <summary>
    /// For the local storage, the credentials are trivially hardcoded.
    /// </summary>
    public (string UserId, string WopiUser) GetUserIdFromCreds(string token, string wopiUser) =>
        ("0:0", "root!0:0");

    public Task<StatInfo> Stat(string endpoint, string filePath, string userId)
    {
        var start = time.GetTimestamp();
        var full = GetFilePath(filePath);
        if (Directory.Exists(full))
        {
            throw new StorageException("Is a directory");
        }

        var info = new FileInfo(full);
        if (!info.Exists)
        {
            throw new StorageException(Common.EnoentMsg);
        }

        var meta = Meta(full);
        var xattrs = ReadMeta(meta);
        log.LogInformation(
            "Invoked stat {inode} {filepath} {elapsedTimems}",
            meta.Inode,
            full,
            Elapsed(start));
        var modified = new DateTimeOffset(info.LastWriteTimeUtc);
        var ticks = modified.UtcTicks - DateTimeOffset.UnixEpoch.UtcTicks;
        var etag = string.Create(CultureInfo.InvariantCulture, $"{ticks / TimeSpan.TicksPerSecond}.{ticks % TimeSpan.TicksPerSecond:D7}");
        return Task.FromResult(
            new StatInfo(
                Common.EncodeInode("local", meta.Inode),
                filePath,
                "0:0",
                info.Length,
                modified.ToUnixTimeSeconds(),
                etag,
                xattrs));
    }

    /// <summary>
    /// Equivalent to <see cref="Stat"/> in the case of local storage.
    /// </summary>
    public Task<StatInfo> StatX(string endpoint, string fileRef, string userId) =>
        Stat(endpoint, fileRef, userId);

    string Elapsed(long start) =>
        time.GetElapsedTime(start).TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture);

    /// <summary>
    /// Common logic for validating locks: duplicates some logic natively implemented by EOS and Reva
    /// on the other storage interfaces.
    /// </summary>
    void ValidateLock(string filePath, StorageLock? currentLock, LockMetadata? lockMetadata, string operation)
    {
        var appName = lockMetadata?.AppName;
        var value = lockMetadata?.LockId;
        try
        {
            if (currentLock == null ||
                currentLock.Expiration.Seconds < Now)
            {
                // overwriting a file with a missing or expired lock is ok,
                // other operations require a valid pre-existing lock
                if (operation == "writefile")
                {
                    return;
                }

                throw new StorageException(Common.ExclError);
            }

            if (!string.IsNullOrEmpty(appName) &&
                currentLock.AppName != appName)
            {
                throw new StorageException($"{Common.ExclError}, file is locked by {currentLock.AppName}");
            }

            if (value != currentLock.LockId)
            {
                throw new StorageException(Common.ExclError);
            }
        }
        catch (StorageException exception)
        {
            log.LogWarning(
                "Failed to {operation} {filepath} {appname} {lockid} {currlock} {reason}",
                operation,
                filePath,
                appName,
                value,
                Json.Serialize(currentLock),
                exception.Message);
            throw;
        }
    }

    public Task SetXattr(string endpoint, string filePath, string userId, string key, string value, LockMetadata? lockMetadata)
    {
        try
        {
            var full = GetExistingFilePath(filePath);
            UpdateMeta(full, _ => _[key] = value);
        }
        catch (StorageException exception)
        {
            log.LogError("Failed to setxattr {filepath} {key} {exception}", filePath, key, exception.Message);
            throw;
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Gets the extended attribute, without failing.
    /// </summary>
    string? GetXattr(string full, string key)
    {
        try
        {
            if (File.Exists(full) &&
                ReadMeta(Meta(full)).TryGetValue(key, out var value))
            {
                return value;
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
        {
            log.LogWarning("Failed to getxattr or missing key {filepath} {key} {exception}", full, key, exception.Message);
        }

        return null;
    }

    public Task RemoveXattr(string endpoint, string filePath, string userId, string key, LockMetadata? lockMetadata)
    {
        try
        {
            var full = GetExistingFilePath(filePath);
            UpdateMeta(
                full,
                _ =>
                {
                    if (!_.Remove(key))
                    {
                        throw new StorageException("No data available");
                    }
                });
        }
        catch (StorageException exception)
        {
            log.LogError("Failed to rmxattr {filepath} {key} {exception}", filePath, key, exception.Message);
            throw;
        }

        return Task.CompletedTask;
    }

    public Task SetLock(string endpoint, string filePath, string userId, string appName, string? value)
    {
        log.LogDebug("Invoked setlock {filepath} {value}", filePath, value);
        var full = GetExistingFilePath(filePath);
        // the metadata update is atomic, so a concurrent setlock either sees this lock or is seen
        UpdateMeta(
            full,
            meta =>
            {
                if (ValidLock(meta) != null)
                {
                    throw new StorageException(Common.ExclError);
                }

                log.LogDebug("setlock: invoking setxattr {filepath} {value}", filePath, value);
                meta[Common.LockKey] = Common.GenRevaLock(appName, value, Now + config.LockExpiration);
            });
        return Task.CompletedTask;
    }

    StorageLock? ValidLock(Dictionary<string, string> meta)
    {
        if (!meta.TryGetValue(Common.LockKey, out var raw))
        {
            return null;
        }

        var storageLock = Common.RetrieveRevaLock(raw);
        if (storageLock.Expiration.Seconds > Now)
        {
            return storageLock;
        }

        return null;
    }

    public Task<StorageLock?> GetLock(string endpoint, string filePath, string userId) =>
        Task.FromResult(GetLock(filePath));

    StorageLock? GetLock(string filePath)
    {
        var full = GetFilePath(filePath);
        var raw = GetXattr(full, Common.LockKey);
        if (raw == null)
        {
            return null;
        }

        var storageLock = Common.RetrieveRevaLock(raw);
        if (storageLock.Expiration.Seconds > Now)
        {
            log.LogDebug("Invoked getlock {filepath}", filePath);
            return storageLock;
        }

        // otherwise, the lock had expired: drop it and return None
        log.LogDebug("getlock: removed stale lock {filepath}", filePath);
        UpdateMeta(full, _ => _.Remove(Common.LockKey));
        return null;
    }

    public Task RefreshLock(string endpoint, string filePath, string userId, string appName, string? value, string? oldValue = null)
    {
        var currentLock = GetLock(filePath);
        if (string.IsNullOrEmpty(oldValue) &&
            currentLock != null)
        {
            // this is a pure refresh operation
            oldValue = currentLock.LockId;
        }

        ValidateLock(filePath, currentLock, new(appName, oldValue), "refreshlock");
        // this is non-atomic, but if we get here the lock was already held
        log.LogDebug("Invoked refreshlock {filepath} {value}", filePath, value);
        UpdateMeta(GetFilePath(filePath), _ => _[Common.LockKey] = Common.GenRevaLock(appName, value, Now + config.LockExpiration));
        return Task.CompletedTask;
    }

    public Task Unlock(string endpoint, string filePath, string userId, string appName, string? value)
    {
        ValidateLock(filePath, GetLock(filePath), new(appName, value), "unlock");
        log.LogDebug("Invoked unlock {filepath} {value}", filePath, value);
        UpdateMeta(GetFilePath(filePath), _ => _.Remove(Common.LockKey));
        return Task.CompletedTask;
    }

    public Task<Stream> ReadFile(string endpoint, string filePath, string userId, string? lockId)
    {
        log.LogDebug("Invoking readFile {filepath}", filePath);
        var start = time.GetTimestamp();
        var full = GetFilePath(filePath);
        if (Directory.Exists(full))
        {
            log.LogError("Error opening the file for read {filepath} {error}", filePath, "Is a directory");
            throw new StorageException("Is a directory");
        }

        try
        {
            var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, ChunkSize, FileOptions.Asynchronous);
            log.LogInformation("File open for read {filepath} {elapsedTimems}", filePath, Elapsed(start));
            return Task.FromResult<Stream>(stream);
        }
        catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException)
        {
            // log this case as info to keep the logs cleaner
            log.LogInformation("File not found on read {filepath}", filePath);
            throw new StorageException(Common.EnoentMsg, exception);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            log.LogError("Error opening the file for read {filepath} {error}", filePath, exception.Message);
            throw ToStorageException(exception);
        }
    }

    /// <summary>
    /// The noVersion argument is ignored for local storage.
    /// </summary>
    public async Task WriteFile(string endpoint, string filePath, string userId, Stream content, long? size, LockMetadata? lockMetadata, bool isLock = false, bool noVersion = false)
    {
        var full = GetFilePath(filePath);
        var currentLock = GetLock(filePath);
        if (lockMetadata != null)
        {
            ValidateLock(filePath, currentLock, lockMetadata, "writefile");
        }
        else if (currentLock != null)
        {
            throw new StorageException(Common.ExclError);
        }

        log.LogDebug("Invoking writeFile {filepath} {size}", filePath, size);
        var start = time.GetTimestamp();
        FileStream file;
        try
        {
            file = new(full, isLock ? FileMode.CreateNew : FileMode.Create, FileAccess.Write, FileShare.None, 1, FileOptions.Asynchronous);
        }
        catch (IOException exception) when (isLock && File.Exists(full))
        {
            log.LogInformation("File exists on write but islock flag requested {filepath}", filePath);
            throw new StorageException(Common.ExclError, exception);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            log.LogError("Error writing file {filepath} {error}", filePath, exception.Message);
            throw ToStorageException(exception);
        }

        var elapsed = Elapsed(start);
        long written;
        await using (file)
        {
            try
            {
                written = await Copy(content, file);
            }
            catch (IOException exception)
            {
                log.LogError("Error writing file {filepath} {error}", filePath, exception.Message);
                throw ToStorageException(exception);
            }
        }

        if (size != null &&
            written != size)
        {
            throw new StorageException($"Written {written} bytes but content is {size} bytes");
        }

        log.LogInformation("File written successfully {filepath} {elapsedTimems} {islock}", filePath, elapsed, isLock);
    }

    async Task<long> Copy(Stream source, Stream target)
    {
        var chunkSize = ChunkSize;
        var buffer = ArrayPool<byte>.Shared.Rent(chunkSize);
        try
        {
            long written = 0;
            int read;
            while ((read = await source.ReadAsync(buffer.AsMemory(0, chunkSize))) > 0)
            {
                await target.WriteAsync(buffer.AsMemory(0, read));
                written += read;
            }

            return written;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    public Task RenameFile(string endpoint, string origFilePath, string newFilePath, string userId, LockMetadata? lockMetadata)
    {
        var currentLock = GetLock(origFilePath);
        if (currentLock != null)
        {
            // enforce lock only if previously set
            ValidateLock(origFilePath, currentLock, lockMetadata, "renamefile");
        }

        var source = GetFilePath(origFilePath);
        var target = GetFilePath(newFilePath);
        try
        {
            File.Move(source, target, overwrite: true);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw ToStorageException(exception);
        }

        MoveMeta(Meta(source), Meta(target));
        return Task.CompletedTask;
    }

    /// <summary>
    /// The force argument is irrelevant and ignored for local storage.
    /// </summary>
    public Task RemoveFile(string endpoint, string filePath, string userId, bool force = false)
    {
        var full = GetExistingFilePath(filePath);
        try
        {
            File.Delete(full);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw ToStorageException(exception);
        }

        DeleteMeta(Meta(full));
        return Task.CompletedTask;
    }

    string GetExistingFilePath(string filePath)
    {
        var full = GetFilePath(filePath);
        if (Directory.Exists(full))
        {
            throw new StorageException("Is a directory");
        }

        if (!File.Exists(full))
        {
            throw new StorageException(Common.EnoentMsg);
        }

        return full;
    }

    static StorageException ToStorageException(Exception exception) =>
        exception switch
        {
            StorageException storageException => storageException,
            FileNotFoundException or DirectoryNotFoundException => new(Common.EnoentMsg, exception),
            UnauthorizedAccessException => new(Common.AccessError, exception),
            _ => new(exception.Message, exception)
        };

    /// <summary>
    /// The metadata document of a file: <see cref="Inode"/> is derived from the path, so it is invariant to
    /// save operations as the WOPI protocol requires.
    /// </summary>
    record MetaFile(string Path, string Inode, object Gate);

    MetaFile Meta(string full)
    {
        var key = System.IO.Path.GetRelativePath(homePath, full).Replace('\\', '/');
        if (pathComparison == StringComparison.OrdinalIgnoreCase)
        {
            key = key.ToUpperInvariant();
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(key));
        var inode = BitConverter.ToUInt64(hash).ToString(CultureInfo.InvariantCulture);
        var path = System.IO.Path.Combine(MetaPath, $"{Convert.ToHexStringLower(hash, 0, 16)}.json");
        return new(path, inode, gates.GetOrAdd(path, _ => new()));
    }

    static Dictionary<string, string> ReadMeta(MetaFile meta)
    {
        lock (meta.Gate)
        {
            if (!File.Exists(meta.Path))
            {
                return [];
            }

            try
            {
                return Retry(
                    () =>
                    {
                        using var stream = new FileStream(meta.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
                        return Deserialize(stream);
                    });
            }
            catch (FileNotFoundException)
            {
                return [];
            }
            catch (Exception exception) when (IsMetaFailure(exception))
            {
                throw ToStorageException(exception);
            }
        }
    }

    void UpdateMeta(string full, Action<Dictionary<string, string>> update)
    {
        var meta = Meta(full);
        lock (meta.Gate)
        {
            try
            {
                CreateMetaFolder();
                bool isEmpty;
                using (var stream = Retry(() => new FileStream(meta.Path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)))
                {
                    var values = Deserialize(stream);
                    update(values);
                    stream.SetLength(0);
                    isEmpty = values.Count == 0;
                    if (!isEmpty)
                    {
                        JsonSerializer.Serialize(stream, values, Json.Options);
                    }
                }

                if (isEmpty)
                {
                    File.Delete(meta.Path);
                }
            }
            catch (Exception exception) when (IsMetaFailure(exception))
            {
                throw ToStorageException(exception);
            }
        }
    }

    static bool IsMetaFailure(Exception exception) =>
        exception is (IOException and not StorageException) or UnauthorizedAccessException or JsonException;

    void CreateMetaFolder()
    {
        if (Directory.Exists(MetaPath))
        {
            return;
        }

        var folder = Directory.CreateDirectory(MetaPath);
        if (OperatingSystem.IsWindows())
        {
            folder.Attributes |= FileAttributes.Hidden;
        }
    }

    static void MoveMeta(MetaFile source, MetaFile target)
    {
        if (source.Path == target.Path)
        {
            return;
        }

        // take the two gates in a stable order to not deadlock with a concurrent opposite move
        var (first, second) = string.CompareOrdinal(source.Path, target.Path) < 0 ? (source, target) : (target, source);
        lock (first.Gate)
        {
            lock (second.Gate)
            {
                if (File.Exists(source.Path))
                {
                    File.Move(source.Path, target.Path, overwrite: true);
                }
                else if (File.Exists(target.Path))
                {
                    File.Delete(target.Path);
                }
            }
        }
    }

    static void DeleteMeta(MetaFile meta)
    {
        lock (meta.Gate)
        {
            if (File.Exists(meta.Path))
            {
                File.Delete(meta.Path);
            }
        }
    }

    static Dictionary<string, string> Deserialize(FileStream stream)
    {
        if (stream.Length == 0)
        {
            return [];
        }

        return JsonSerializer.Deserialize<Dictionary<string, string>>(stream, Json.Options) ?? [];
    }

    /// <summary>
    /// Metadata documents are opened exclusively for updates, so another process may hold one briefly.
    /// </summary>
    static T Retry<T>(Func<T> action)
    {
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                return action();
            }
            catch (IOException exception) when (attempt < 100 && exception is not (FileNotFoundException or DirectoryNotFoundException))
            {
                Thread.Sleep(10);
            }
        }
    }
}
