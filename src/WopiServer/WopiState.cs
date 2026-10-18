/// <summary>
/// All state information of the WOPI server: the settings read at startup, plus the bookkeeping of open
/// files and conflicted sessions, which is only exposed to operators.
/// </summary>
public class WopiState
{
    public const string Version = "git";

    public required WopiConfig Config { get; init; }
    public required IStorage Storage { get; init; }
    public required TimeProvider Time { get; init; }
    public required string StorageType { get; init; }
    public required int Port { get; init; }
    public required bool UseHttps { get; init; }
    public required string WopiSecret { get; init; }
    public required string IopSecret { get; init; }
    public required string WopiUrl { get; init; }
    public required string HomePath { get; init; }
    public required string RecoveryPath { get; init; }
    public required string WopiProxy { get; init; }
    public required string? WopiProxyKey { get; init; }
    public required string ProxiedAppName { get; init; }
    public required IReadOnlyList<string> NonOfficeTypes { get; init; }
    public required IReadOnlyList<string> CodeTypes { get; init; }
    public required bool EnableRename { get; init; }

    /// <summary>
    /// The validation of X-WOPI-Proof, when configured.
    /// </summary>
    internal ProofKeys? ProofKeys { get; init; }

    public long Now => Time.GetUtcNow().ToUnixTimeSeconds();

    readonly Lock sync = new();

    // filename -> (time first seen, sessions)
    readonly Dictionary<string, (string Time, HashSet<string> Sessions)> openFiles = [];

    // sessions for which a lock conflict is outstanding or resolved
    readonly Dictionary<string, PendingSession> pending = [];
    readonly Dictionary<string, ResolvedSession> resolved = [];
    readonly HashSet<string> allUsers = [];

    public void AddUser(string userId)
    {
        lock (sync)
        {
            allUsers.Add(userId);
        }
    }

    public bool IsPending(string? session, out long time)
    {
        lock (sync)
        {
            if (session != null &&
                pending.TryGetValue(session, out var value))
            {
                time = value.Time;
                return true;
            }

            time = 0;
            return false;
        }
    }

    /// <summary>
    /// Records a conflicted session, unless the session is already known.
    /// </summary>
    public void AddPending(string session, string user, string? heldBy, string type)
    {
        lock (sync)
        {
            if (pending.ContainsKey(session) ||
                resolved.ContainsKey(session))
            {
                return;
            }

            pending[session] = new(user, Now, heldBy, type);
        }
    }

    /// <summary>
    /// Marks a session as resolved and accounts the given filename in the open files map.
    /// This is only used for bookkeeping, no functionality is associated to those maps.
    /// </summary>
    public void ResolveSession(string? session, string filename)
    {
        lock (sync)
        {
            if (session != null &&
                pending.Remove(session, out var value))
            {
                resolved[session] = new(value.User, Now - value.Time, value.Type);
            }

            // keep some accounting of the open files
            if (!openFiles.TryGetValue(filename, out var entry))
            {
                entry = (DateTime.Now.ToString("ddd MMM d HH:mm:ss yyyy", CultureInfo.InvariantCulture), []);
                openFiles[filename] = entry;
            }

            entry.Sessions.Add(session ?? "None");
        }
    }

    /// <summary>
    /// On unlock, the file is closed and a pending conflict on the session is resolved.
    /// </summary>
    public void CloseFile(string filename, string? session)
    {
        lock (sync)
        {
            if (!openFiles.Remove(filename))
            {
                // already removed?
                return;
            }

            if (session != null &&
                pending.Remove(session, out var value))
            {
                resolved[session] = new(value.User, Now - value.Time, null);
            }
        }
    }

    public string OpenFilesJson()
    {
        lock (sync)
        {
            var list = new JsonObject();
            foreach (var (filename, (time, sessions)) in openFiles)
            {
                list[filename] = new JsonArray(time, new JsonArray(sessions.Select(_ => JsonValue.Create(_)).ToArray<JsonNode?>()));
            }

            return list.ToJsonString(Json.Options);
        }
    }

    public string ConflictsJson()
    {
        lock (sync)
        {
            return Json.Serialize(
                new
                {
                    pending,
                    resolved,
                    users = allUsers.Count
                });
        }
    }
}

public record PendingSession(
    [property: JsonPropertyName("user")] string User,
    [property: JsonPropertyName("time")] long Time,
    [property: JsonPropertyName("heldby")] string? HeldBy,
    [property: JsonPropertyName("type")] string Type);

public record ResolvedSession(
    [property: JsonPropertyName("user")] string User,
    [property: JsonPropertyName("restime")] long ResTime,
    [property: JsonPropertyName("type")] string? Type);
