/// <summary>
/// A client in the file coauthoring tracker ([MS-FSSHTTP] 3.1.1), sharing the lock until it expires.
/// </summary>
record SessionClient(string ClientId, DateTimeOffset Expires, int TimeoutSeconds);

/// <summary>
/// An entry of the editors table ([MS-FSSHTTPB] 2.2.3.1.2.1.2).
/// </summary>
record SessionEditor(string ClientId, string UserId, string? UserName, bool AsEditor, DateTimeOffset Expires, Dictionary<string, string> Metadata);

/// <summary>
/// The file coauthoring tracker of a file ([MS-FSSHTTP] 3.1.1): the shared lock (its schema lock id and the clients that
/// share it), or the exclusive lock, whether a coauthoring transition is pending (3.1.4.3.1), and the editors table.
/// </summary>
record Session(
    string? SchemaLockId,
    IReadOnlyList<SessionClient> Clients,
    string? ExclusiveLockId,
    DateTimeOffset ExclusiveExpires,
    bool TransitionPending,
    IReadOnlyList<SessionEditor> Editors)
{
    public static Session None { get; } = new(null, [], null, default, false, []);

    /// <summary>
    /// Without the clients, editors and exclusive lock that expired; the shared lock goes with its last client.
    /// </summary>
    public Session Purge(DateTimeOffset now)
    {
        var clients = Clients.Where(_ => _.Expires > now).ToList();
        var exclusive = ExclusiveLockId;
        if (ExclusiveExpires <= now)
        {
            exclusive = null;
        }

        var schema = SchemaLockId;
        var pending = TransitionPending;
        if (clients.Count == 0)
        {
            schema = null;
        }

        // a transition to coauthoring is over when the client is alone again
        if (clients.Count <= 1)
        {
            pending = false;
        }

        return new(schema, clients, exclusive, ExclusiveExpires, pending, Editors.Where(_ => _.Expires > now).ToList());
    }

    public SessionClient? Client(string clientId) => Clients.FirstOrDefault(_ => _.ClientId == clientId);

    /// <summary>
    /// The <see cref="OfficeLock"/> that stands for the lock of the session in the <see cref="IOfficeLockStore"/>, so
    /// that WOPI clients see the file locked: null when there is none.
    /// </summary>
    public OfficeLock? Lock()
    {
        if (ExclusiveLockId != null)
        {
            return new(CoauthSessions.ExclusivePrefix + ExclusiveLockId, ExclusiveExpires);
        }

        if (SchemaLockId != null &&
            Clients.Count > 0)
        {
            return new(CoauthSessions.SchemaPrefix + SchemaLockId, Clients.Max(_ => _.Expires));
        }

        return null;
    }
}

/// <summary>
/// The outcome of an operation on a session: the error code of the sub-response, and the session after it.
/// </summary>
record SessionResult(string ErrorCode, Session Session);

/// <summary>
/// The file coauthoring trackers, kept in the <see cref="IFsshttpStore"/> under a partition of their own, with their lock
/// in the <see cref="IOfficeLockStore"/>: a WOPI lock blocks desktop coauthoring, and desktop coauthoring blocks WOPI
/// locks, without either knowing of the other. An operation reads the tracker, drops what expired, changes it, takes or
/// releases the lock, and stores the tracker only if no other operation stored it meanwhile; otherwise it starts again.
/// </summary>
class CoauthSessions(IFsshttpStore store, IOfficeLockStore locks, TimeProvider time)
{
    public const string SchemaPrefix = "fsshttp-schema:";
    public const string ExclusivePrefix = "fsshttp-exclusive:";

    /// <summary>
    /// The partition of the tracker in the <see cref="IFsshttpStore"/>, which no Office partition uses.
    /// </summary>
    public static Guid SessionPartition { get; } = new("B5D0E8A3-3C51-4F6B-9E2A-6C1D7F4A2E90");

    const int maxAttempts = 8;

    static JsonSerializerOptions json = new(JsonSerializerDefaults.Web);

    public DateTimeOffset Now => time.GetUtcNow();

    public async Task<Session> Get(string fileId, Cancel cancel = default) =>
        (await Read(fileId, cancel)).Session.Purge(Now);

    async Task<(string? Version, Session Session)> Read(string fileId, Cancel cancel)
    {
        var stored = await store.Get(fileId, SessionPartition, cancel);
        if (stored == null)
        {
            return (null, Session.None);
        }

        var session = JsonSerializer.Deserialize<Session>(stored.Data.Span, json) ?? Session.None;
        return (stored.Version, session);
    }

    /// <summary>
    /// Changes the session of the file. <paramref name="change"/> gets the current session, without what expired, and
    /// returns the error code and the new session; the session and the lock are written only on "Success". A lock in the
    /// <see cref="IOfficeLockStore"/> that the session does not hold (a WOPI lock) fails with "FileAlreadyLockedOnServer"
    /// when the new session needs one.
    /// </summary>
    public async Task<SessionResult> Change(string fileId, Func<Session, SessionResult> change, Cancel cancel = default)
    {
        for (var attempt = 0; attempt < maxAttempts; attempt++)
        {
            var (version, stored) = await Read(fileId, cancel);
            var current = stored.Purge(Now);
            var result = change(current);
            if (result.ErrorCode != "Success")
            {
                return result;
            }

            var wanted = result.Session.Lock();
            var existing = await locks.Get(fileId, cancel);
            var foreign = existing != null && !IsSessionLock(existing.Id);
            if (foreign)
            {
                // a WOPI lock: the session cannot take one, but can let go of its own
                if (wanted != null)
                {
                    return new("FileAlreadyLockedOnServer", current);
                }
            }
            else if (existing?.Id != wanted?.Id ||
                     existing?.Expires != wanted?.Expires)
            {
                var previous = await locks.CompareExchange(fileId, existing?.Id, wanted, cancel);
                if (previous?.Id != existing?.Id)
                {
                    continue;
                }
            }

            var data = JsonSerializer.SerializeToUtf8Bytes(result.Session, json);
            if (await store.CompareExchange(fileId, SessionPartition, version, new(Guid.NewGuid().ToString("N"), data), cancel))
            {
                return result;
            }
        }

        return new("CoauthRefblobConcurrencyViolation", Session.None);
    }

    public static bool IsSessionLock(string lockId) =>
        lockId.StartsWith(SchemaPrefix, StringComparison.Ordinal) ||
        lockId.StartsWith(ExclusivePrefix, StringComparison.Ordinal);

    /// <summary>
    /// The timeout of a client in seconds: at least an hour, at most 120,000 seconds ([MS-FSSHTTP] 2.3.1.5: from 60 to
    /// 3,600 the server picks its own).
    /// </summary>
    public static int Timeout(int? seconds) => Math.Clamp(seconds ?? 3600, 3600, 120_000);

    /// <summary>
    /// Whether a lock id given with a Put Changes (BypassLockID) is the one the file holds, or the file holds none.
    /// </summary>
    public async Task<bool> CanWrite(string fileId, string? lockId, Cancel cancel = default)
    {
        var existing = await locks.Get(fileId, cancel);
        if (existing == null)
        {
            return true;
        }

        return lockId != null &&
               (existing.Id == SchemaPrefix + lockId ||
                existing.Id == ExclusivePrefix + lockId);
    }
}
