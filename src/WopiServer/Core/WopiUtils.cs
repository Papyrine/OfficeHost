/// <summary>
/// General low-level functions to support the WOPI server.
/// </summary>
class WopiUtils(WopiState state, ILogger log)
{
    /// <summary>
    /// The xattr key used for conflicts resolution on the remote storage,
    /// as well as to establish whether a new version is to be created on save.
    /// </summary>
    public const string LastSaveTimeKey = "iop.wopi.lastwritetime";

    /// <summary>
    /// The xattr key used to store user info data from WOPI apps.
    /// </summary>
    public const string UserInfoKey = "iop.wopi.userinfo";

    /// <summary>
    /// Conventional string representing an external, non-WOPI lock.
    /// </summary>
    public const string ExternalLock = "External";

    IStorage Storage => state.Storage;

    WopiConfig Config => state.Config;

    /// <summary>
    /// Validates the access_token, and logs some additional headers of the request.
    /// </summary>
    public (AccessToken? Token, WopiResult? Error) ValidateAndLogHeaders(HttpContext context, string operation)
    {
        var request = context.Request;
        AccessToken acctok;
        try
        {
            acctok = Jwt.Decode<AccessToken>(request.Arg("access_token") ?? throw new JwtException("'access_token'"), state.WopiSecret, state.Now);
            if (acctok.Exp < state.Now ||
                !acctok.Iss.Contains("cs3org:wopiserver"))
            {
                throw new JwtException("Signature has expired", expired: true);
            }
        }
        catch (JwtException exception)
        {
            log.LogInformation(
                "Expired or malformed token {client} {requestedUrl} {details} {token}",
                context.Client(),
                request.BaseUrl(),
                exception.Message,
                request.Arg("access_token"));
            return (null, WopiResult.Text("Invalid access token", StatusCodes.Status401Unauthorized));
        }

        // validate the WOPI timestamp: this is typically not present, but if it is we must check its expiration
        // (cf. the WOPI validator tests)
        var wopiTimestamp = "None";
        var timestampHeader = request.Header("X-WOPI-TimeStamp");
        if (timestampHeader != null)
        {
            // the timestamp is in .NET ticks; timestamps older than 20 minutes must be considered expired
            if (!long.TryParse(timestampHeader, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ticks) ||
                ticks < state.Time.GetUtcNow().UtcTicks - TimeSpan.FromMinutes(20).Ticks)
            {
                log.LogWarning(
                    "{operation}: invalid X-WOPI-Timestamp {user} {token} {client}",
                    operation,
                    acctok.UserId.Short(),
                    context.TokenForLog(),
                    context.Client());
                // UNAUTHORIZED would seem more appropriate here, but the ProofKeys part of the MS test suite explicitly requires this
                return (null, WopiResult.Text("Invalid or expired X-WOPI-Timestamp header", StatusCodes.Status500InternalServerError));
            }

            wopiTimestamp = (ticks / 10000000.0).ToString(CultureInfo.InvariantCulture);
        }

        // log all relevant headers to help debugging
        var session = request.Header("X-WOPI-SessionId");
        log.LogInformation(
            "{operation}: client context {trace} {user} {filename} {token} {client} {deviceId} {reqId} {sessionId} {app} {appEndpoint} {appVersion} {correlationId} {wopits}",
            operation.Title(),
            acctok.Trace ?? "N/A",
            acctok.UserId.Short(),
            acctok.FileName,
            context.TokenForLog(),
            context.Client(),
            request.Header("X-WOPI-DeviceId"),
            request.Header("X-Request-Id"),
            session,
            request.Header("X-WOPI-RequestingApplication"),
            request.Header("X-WOPI-AppEndpoint"),
            request.Header("X-WOPI-ClientVersion"),
            request.Header("X-WOPI-CorrelationId"),
            wopiTimestamp);

        // update bookkeeping of pending sessions
        if (operation.Title() == "Checkfileinfo" &&
            state.IsPending(session, out var time) &&
            time < state.Now - 30)
        {
            // a previously conflicted session is still around executing Checkfileinfo after some time, assume it got resolved
            state.ResolveSession(session, acctok.FileName);
        }

        return (acctok, null);
    }

    /// <summary>
    /// Returns a URL-encoded WOPISrc for the given fileid, proxied if required.
    /// </summary>
    public string GenerateWopiSrc(string fileId, bool proxy = false)
    {
        if (!proxy ||
            state.WopiProxy.Length == 0)
        {
            return UrlCodec.QuotePlus($"{state.WopiUrl}/wopi/files/{fileId}").Replace("-", "%2D");
        }

        // proxy the WOPI request through an external WOPI proxy service, but only if it was not already proxied
        // heuristically, proxied fileids are (much) longer than that
        if (fileId.Length < 90)
        {
            log.LogDebug("Generating proxied fileid {fileid} {proxy}", fileId, state.WopiProxy);
            fileId = Jwt.Encode(
                new JsonObject
                {
                    ["u"] = $"{state.WopiUrl}/wopi/files/",
                    ["f"] = fileId
                },
                state.WopiProxyKey!);
        }
        else
        {
            log.LogDebug("Proxied fileid already created {fileid} {proxy}", fileId, state.WopiProxy);
        }

        return UrlCodec.QuotePlus($"{state.WopiProxy}/wopi/files/{fileId}").Replace("-", "%2D");
    }

    /// <summary>
    /// Parses a URL template and returns it with actualised placeholders.
    /// </summary>
    public static string GenerateUrlFromTemplate(string url, string userName, UserType userType, string folderUrl, string filename, string endpoint, string fileId, string appName)
    {
        string path;
        if (userName.Length == 0 ||
            userType == UserType.Anonymous)
        {
            var folderPath = Uri.TryCreate(folderUrl, UriKind.Absolute, out var uri) ? uri.AbsolutePath : folderUrl.Split('?', '#')[0];
            path = $"{folderPath.Replace("/files/link/", "/")}/{PosixPath.BaseName(filename)}";
        }
        else
        {
            path = filename;
        }

        return url
            .Replace("<path>", path)
            .Replace("<endpoint>", endpoint)
            .Replace("<fileid>", fileId)
            .Replace("<app>", appName);
    }

    public static string GenerateUrlFromTemplate(string url, AccessToken acctok) =>
        GenerateUrlFromTemplate(url, acctok.UserName, acctok.UserType, acctok.FolderUrl, acctok.FileName, acctok.Endpoint, acctok.FileId, acctok.AppName);

    /// <summary>
    /// The filename of a LibreOffice-compatible lock file.
    /// This enables interoperability between Online and Desktop applications.
    /// </summary>
    public static string GetLibreOfficeLockName(string filename) =>
        PosixPath.Join(PosixPath.DirName(filename), $".~lock.{PosixPath.BaseName(filename)}#");

    /// <summary>
    /// The filename of a lock file as created by Microsoft Office.
    /// </summary>
    public static string GetMicrosoftOfficeLockName(string filename)
    {
        var dir = PosixPath.DirName(filename);
        var name = PosixPath.BaseName(filename);
        if (PosixPath.Extension(filename) != ".docx" ||
            name.Length <= 6 + 1 + 4)
        {
            return PosixPath.Join(dir, $"~${name}");
        }

        // MS Word has a really weird algorithm for the lock file name...
        if (name.Length >= 8 + 1 + 4)
        {
            return PosixPath.Join(dir, $"~${name[2..]}");
        }

        return PosixPath.Join(dir, $"~${name[1..]}");
    }

    public static string RandomString(int size) =>
        new(Random.Shared.GetItems<char>("abcdefghijklmnopqrstuvwxyz", size));

    /// <summary>
    /// Generates an access token for a given file and a given user.
    /// </summary>
    /// <param name="wopiUser">In the form `username!userid_in_stat_format`.</param>
    public async Task<GeneratedToken> GenerateAccessToken(
        string userId,
        string fileId,
        ViewMode viewMode,
        string friendlyName,
        string wopiUser,
        UserType userType,
        string folderUrl,
        string endpoint,
        string appName,
        string? appEditUrl,
        string? appViewUrl,
        string trace)
    {
        log.LogDebug("Generating token {userid} {fileid} {endpoint} {app}", userId.Short(), fileId, endpoint, appName);
        StatInfo statInfo;
        try
        {
            // stat the file to check for existence and get a version-invariant inode:
            // the inode serves as fileid (and must not change across save operations)
            statInfo = await Storage.StatX(endpoint, fileId, userId);
        }
        catch (StorageException exception)
        {
            log.LogInformation("Requested file not found or not a file {fileid} {error}", fileId, exception.Message);
            throw;
        }

        var expTime = state.Now + Config.GetInt("general", "tokenvalidity", 86400);
        var filename = statInfo.FilePath;
        var extension = PosixPath.Extension(filename).ToLowerInvariant();

        var details = "";
        string[] legacyFormats = [".doc", ".dot", ".xls", ".ppt", ".pps", ".csv"];
        if (!appName.Contains("Collabora") &&
            viewMode == ViewMode.ReadWrite &&
            (legacyFormats.Contains(extension) ||
             IsOpenDocument(extension) && Config.IsTrue("general", "disablemswriteodf", false)))
        {
            // we're opening a legacy format file or an ODF (`.o[d|t]?`) and the app is not Collabora
            log.LogInformation("Forcing read-only access to ODF/legacy formats {filename}", filename);
            viewMode = ViewMode.ReadOnly;
            details = "App does not support edit for this file type";
        }

        // if we're opening the file in edit mode, check for existing locks and downgrade to read-only if any is found
        if (viewMode.CanWrite() &&
            Config.IsTrue("general", "detectexternallocks", true))
        {
            var (retrievedLock, holder) = await RetrieveWopiLock(fileId, "genAccessToken", "NA", endpoint, filename, userId, "NA");
            if (retrievedLock == ExternalLock ||
                holder != null && holder != appName)
            {
                log.LogInformation("Forcing read-only access due to existing lock {filename} {lock} {app}", filename, $"({retrievedLock}, {holder})", appName);
                viewMode = ViewMode.ReadOnly;
                if (retrievedLock == ExternalLock)
                {
                    details = $"This file is opened exclusively by {holder} and cannot be edited";
                }
                else
                {
                    details = $"This file cannot be edited in {appName}, use app {holder} to edit instead";
                }
            }
        }

        if (viewMode == ViewMode.Preview &&
            statInfo.Size == 0)
        {
            // override preview mode when a new file is being created. No need to provide a reason
            viewMode = ViewMode.ReadWrite;
        }

        var includeAppUrls = !Config.Has("general", "hostediturl") ||
                             !Config.Has("general", "hostviewurl");
        var payload = new AccessToken
        {
            UserId = userId,
            WopiUser = wopiUser,
            UserType = userType,
            FileName = filename,
            FileId = fileId,
            UserName = friendlyName,
            ViewMode = viewMode,
            FolderUrl = folderUrl,
            Endpoint = endpoint,
            AppName = appName,
            Trace = trace,
            Exp = expTime,
            Iss = $"cs3org:wopiserver:{WopiState.Version}",
            AppEditUrl = includeAppUrls ? appEditUrl : null,
            AppViewUrl = includeAppUrls ? appViewUrl : null
        };
        var acctok = Jwt.Encode(payload, state.WopiSecret);
        state.AddUser(userId);

        log.LogInformation(
            "Access token generated {trace} {userid} {wopiuser} {friendlyname} {usertype} {mode} {endpoint} {filename} {inode} {mtime} {folderurl} {appname} {expiration} {token}",
            trace,
            userId.Short(),
            wopiUser,
            friendlyName,
            userType.Value(),
            viewMode.Value(),
            endpoint,
            filename,
            statInfo.Inode,
            statInfo.MTime,
            folderUrl,
            appName,
            expTime,
            acctok.Short());
        return new(statInfo.Inode, acctok, viewMode, details);
    }

    static bool IsOpenDocument(string extension) =>
        extension.Length >= 3 &&
        extension[1..3] is "od" or "ot";

    /// <summary>
    /// Generates the lock payload for the storage given the raw metadata.
    /// </summary>
    public static string? EncodeLock(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        return $"{Common.WebDavLockPrefix} {Convert.ToBase64String(Encoding.UTF8.GetBytes(value))}";
    }

    /// <summary>
    /// Restores the lock payload reverting the <see cref="EncodeLock"/> format.
    /// </summary>
    static string DecodeLock(string? storedLock)
    {
        if (storedLock == null ||
            !storedLock.StartsWith(Common.WebDavLockPrefix, StringComparison.Ordinal))
        {
            // it's not our lock, though it's likely valid
            throw new StorageException("Non-WOPI lock found");
        }

        try
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(storedLock[(Common.WebDavLockPrefix.Length + 1)..]));
        }
        catch (FormatException exception)
        {
            throw new StorageException(exception.Message, exception);
        }
    }

    /// <summary>
    /// Retrieves and logs a lock for a given file: returns the lock and its holder, or (null, null) if no lock found.
    /// </summary>
    public async Task<(string? Lock, string? Holder)> RetrieveWopiLock(
        string fileId,
        string operation,
        string? lockForLog,
        string endpoint,
        string filename,
        string userId,
        string tokenForLog,
        string? overrideFilename = null)
    {
        // if required, check if a non-WOPI office lock exists for this file
        var checkExternal = Config.IsTrue("general", "detectexternallocks", true);
        string? libreOfficeLock = null;
        if (checkExternal &&
            !state.NonOfficeTypes.Contains(PosixPath.Extension(filename)))
        {
            try
            {
                // first try to look for a MS Office lock
                var msLockStat = await Storage.Stat(endpoint, GetMicrosoftOfficeLockName(filename), userId);
                log.LogInformation(
                    "Found existing MS Office lock {lockop} {user} {filename} {token} {lockmtime}",
                    operation.Title(),
                    userId.Short(),
                    filename,
                    tokenForLog,
                    msLockStat.MTime);
                return (ExternalLock, $"{msLockStat.OwnerId.Split('@')[0]} via Microsoft Office for Desktop");
            }
            catch (StorageException)
            {
            }

            try
            {
                // then try to read a LibreOffice lock
                var lockName = GetLibreOfficeLockName(filename);
                var loLockStat = await Storage.StatX(endpoint, lockName, userId);
                libreOfficeLock = await ReadText(endpoint, lockName, userId);
                if (!libreOfficeLock.Contains("wopiserver"))
                {
                    var holder = libreOfficeLock.Contains(',') ? libreOfficeLock.Split(',')[1] : loLockStat.OwnerId;
                    log.LogInformation(
                        "Found existing LibreOffice lock {lockop} {user} {filename} {token} {lockmtime} {holder}",
                        operation.Title(),
                        userId.Short(),
                        filename,
                        tokenForLog,
                        loLockStat.MTime,
                        holder);
                    return (ExternalLock, $"{holder} via LibreOffice for Desktop");
                }
            }
            catch (StorageException)
            {
            }
        }

        StorageLock? lockContent;
        string decodedLock;
        try
        {
            // fetch and decode the lock
            lockContent = await Storage.GetLock(endpoint, overrideFilename ?? filename, userId);
            // here we used to check the last save time to "extend" the validity of an otherwise expired lock:
            // however, this goes against isolating the lock expiration logic in the storage interfaces and ultimately
            // violates the WOPI specifications, therefore it was dropped
            if (lockContent == null)
            {
                log.LogInformation("No lock found {lockop} {user} {filename} {token}", operation.Title(), userId.Short(), filename, tokenForLog);
                // lazily remove the LibreOffice-compatible lock file, if it was detected and has
                // the expected signature - cf. SetLock()
                try
                {
                    if (!string.IsNullOrEmpty(libreOfficeLock))
                    {
                        await Storage.RemoveFile(endpoint, GetLibreOfficeLockName(filename), userId, force: true);
                    }
                }
                catch (StorageException exception)
                {
                    log.LogWarning(
                        "Unable to delete stale LibreOffice-compatible lock file {lockop} {user} {filename} {fileid} {error}",
                        operation.Title(),
                        userId.Short(),
                        filename,
                        fileId,
                        exception.Message);
                }

                return (null, null);
            }

            decodedLock = DecodeLock(lockContent.LockId);
        }
        catch (StorageException exception)
        {
            log.LogInformation(
                "Found non-compatible or unreadable lock {lockop} {user} {filename} {token} {error}",
                operation.Title(),
                userId.Short(),
                filename,
                tokenForLog,
                exception.Message);
            return (ExternalLock, "another app or user");
        }

        log.LogInformation(
            "Retrieved lock {lockop} {user} {filename} {fileid} {lock} {retrievedlock} {expTime} {token}",
            operation.Title(),
            userId.Short(),
            filename,
            fileId,
            lockForLog,
            decodedLock,
            DateTimeOffset.FromUnixTimeSeconds(lockContent.Expiration.Seconds).ToLocalTime().ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
            tokenForLog);
        return (decodedLock, lockContent.AppName);
    }

    /// <summary>
    /// Reads a small file, typically a lock file, as text.
    /// </summary>
    public async Task<string> ReadText(string endpoint, string filename, string userId)
    {
        await using var stream = await Storage.ReadFile(endpoint, filename, userId, null);
        using var reader = new StreamReader(stream, Encoding.UTF8);
        try
        {
            return await reader.ReadToEndAsync();
        }
        catch (IOException exception) when (exception is not StorageException)
        {
            throw new StorageException(exception.Message, exception);
        }
    }

    /// <summary>
    /// Compares two locks and returns true if they represent the same WOPI lock.
    /// Officially, the comparison must be based on the locks' string representations. But because of
    /// a bug in early versions of Word Online, the internal format of the WOPI locks may be looked at,
    /// based on heuristics. Note that this format is subject to change and is not documented!
    /// </summary>
    public bool CompareWopiLocks(string? lock1, string? lock2)
    {
        if (lock1 == lock2)
        {
            log.LogDebug("compareLocks {lock1} {lock2} {result}", lock1, lock2, "True");
            return true;
        }

        if (Config.IsTrue("general", "wopilockstrictcheck", true))
        {
            log.LogDebug("compareLocks {lock1} {lock2} {strict} {result}", lock1, lock2, "True", "False");
            return false;
        }

        // before giving up, attempt to parse the lock as a JSON dictionary if allowed by the config
        var result = CompareRelaxed(lock1, lock2);
        log.LogDebug("compareLocks {lock1} {lock2} {strict} {result}", lock1, lock2, "False", result ? "True" : "False");
        return result;
    }

    static bool CompareRelaxed(string? lock1, string? lock2)
    {
        var l1 = ParseJson(lock1);
        if (l1 is not JsonObject o1 ||
            !o1.TryGetPropertyValue("S", out var s1))
        {
            // lock1 is not a JSON dictionary: fail the comparison
            return false;
        }

        if (ParseJson(lock2) is JsonObject o2)
        {
            // used by Word
            return o2.TryGetPropertyValue("S", out var s2) &&
                   JsonNode.DeepEquals(s1, s2);
        }

        // lock2 is not a JSON dictionary: also used by Word
        return s1 is JsonValue value &&
               value.TryGetValue<string>(out var text) &&
               text == lock2;
    }

    static JsonNode? ParseJson(string? value)
    {
        if (value == null)
        {
            return null;
        }

        try
        {
            return JsonNode.Parse(value);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>
    /// Generates and logs an HTTP 409 response in case of locks conflict.
    /// </summary>
    public WopiResult MakeConflictResponse(HttpContext context, string operation, string user, string? retrievedLock, string? requestedLock, string? oldLock, string filename, string reason, string? saveTime = null) =>
        MakeConflictResponse(
            context,
            operation,
            user,
            retrievedLock,
            requestedLock,
            oldLock,
            filename,
            new JsonObject
            {
                ["message"] = reason
            },
            saveTime);

    public WopiResult MakeConflictResponse(HttpContext context, string operation, string user, string? retrievedLock, string? requestedLock, string? oldLock, string filename, JsonObject reason, string? saveTime = null)
    {
        var message = (string?) reason["message"] ?? "";
        var response = WopiResult.Json(reason, StatusCodes.Status409Conflict);
        response.Headers["X-WOPI-Lock"] = retrievedLock ?? "";
        response.Headers["X-WOPI-LockFailureReason"] = message;
        if (retrievedLock is ExternalLock or "External App")
        {
            // the file is locked by a desktop application, cf. [MS-WOPI] 3.3.5.1.14
            response.Headers["X-WOPI-LockedByOtherInterface"] = "true";
        }

        var session = context.Request.Header("X-WOPI-SessionId");
        if (!string.IsNullOrEmpty(session) &&
            retrievedLock != ExternalLock)
        {
            state.AddPending(session, user, retrievedLock, PosixPath.Extension(filename));
        }

        var fileAge = "NA";
        if (long.TryParse(saveTime, CultureInfo.InvariantCulture, out var saved))
        {
            fileAge = (state.Time.GetUtcNow().ToUnixTimeMilliseconds() / 1000.0 - saved).ToString("F1", CultureInfo.InvariantCulture);
        }

        log.LogWarning(
            "Returning conflict {lockop} {user} {filename} {token} {sessionId} {lock} {oldlock} {retrievedlock} {fileage} {reason}",
            !string.IsNullOrEmpty(oldLock) && oldLock != "NA" && operation != "PUTFILE" ? "UnlockAndRelock" : operation.Title(),
            user.Short(),
            filename,
            context.TokenForLog(),
            session,
            requestedLock,
            oldLock,
            retrievedLock,
            fileAge,
            message);
        return response;
    }

    /// <summary>
    /// Generates and logs an HTTP 200 response with appropriate headers for Lock/RefreshLock operations.
    /// </summary>
    public WopiResult MakeLockSuccessResponse(HttpContext context, string operation, AccessToken acctok, string? requestedLock, string? oldLock, string version)
    {
        var session = context.Request.Header("X-WOPI-SessionId");
        if (string.IsNullOrEmpty(session))
        {
            session = acctok.WopiUserName;
        }

        state.ResolveSession(session, acctok.FileName);

        log.LogInformation(
            "Successfully locked {lockop} {filename} {token} {sessionId} {lock} {oldlock} {version}",
            string.IsNullOrEmpty(oldLock) ? operation.Title() : "UnlockAndRelock",
            acctok.FileName,
            context.TokenForLog(),
            session,
            requestedLock,
            oldLock,
            version);
        var response = new WopiResult(StatusCodes.Status200OK);
        response.Headers["X-WOPI-ItemVersion"] = version;
        return response;
    }

    /// <summary>
    /// Saves the body of the request to the given target filename (defaulting to the access token's one),
    /// and stores the save time as an xattr. Fails with a <see cref="StorageException"/>.
    /// </summary>
    public async Task StoreWopiFile(HttpContext context, AccessToken acctok, string? retrievedLock, string xaKey, string targetName = "", bool noVersion = false)
    {
        if (targetName.Length == 0)
        {
            targetName = acctok.FileName;
        }

        var session = context.Request.Header("X-WOPI-SessionId");
        if (string.IsNullOrEmpty(session))
        {
            session = acctok.WopiUserName;
        }

        state.ResolveSession(session, targetName);

        var lockMetadata = new LockMetadata(acctok.AppName, EncodeLock(retrievedLock));
        StorageException? writeError = null;
        try
        {
            if (context.Request.Body.CanSeek)
            {
                context.Request.Body.Position = 0;
            }

            await Storage.WriteFile(acctok.Endpoint, targetName, acctok.UserId, context.Request.Body, context.Request.ContentLength, lockMetadata, noVersion: noVersion);
        }
        catch (StorageException exception)
        {
            if (exception.IsAccessDenied)
            {
                throw;
            }

            // something went wrong on write: we still want to setxattr but report this error to the caller
            writeError = exception;
        }

        // in all cases save the current time for later conflict checking: this is never older than the mtime of the file
        await Storage.SetXattr(acctok.Endpoint, targetName, acctok.UserId, xaKey, state.Now.ToString(CultureInfo.InvariantCulture), lockMetadata);
        if (writeError != null)
        {
            throw writeError;
        }
    }

    /// <summary>
    /// Saves a conflict file in case the original file was externally locked or overwritten.
    /// The conflicted copy follows the format `&lt;filename&gt;-webconflict-&lt;time&gt;` and might be stored
    /// next to the original one, or to the user's home, or to the recovery path.
    /// </summary>
    public async Task<WopiResult> StoreAfterConflict(HttpContext context, AccessToken acctok, string? retrievedLock, string? requestedLock, string reason)
    {
        var (name, extension) = PosixPath.SplitExt(acctok.FileName);
        // typical EFSS formats are like '<filename>_conflict-<date>-<time>', but they're not synchronized: use a similar format
        var newName = $"{name}-webconflict-{DateTime.Now.ToString("yyyyMMdd-HH", CultureInfo.InvariantCulture)}{extension.Trim()}";
        StorageException? recoveryError = null;
        try
        {
            await StoreWopiFile(context, acctok, retrievedLock, LastSaveTimeKey, newName);
        }
        catch (StorageException exception)
        {
            if (!exception.Message.Contains(Common.AccessError))
            {
                recoveryError = exception;
            }
            else
            {
                // let's try the configured user's (or owner's) homepath instead of the current folder
                newName = PosixPath.Join(HomePathFor(acctok), PosixPath.BaseName(newName));
                try
                {
                    await StoreWopiFile(context, acctok, retrievedLock, LastSaveTimeKey, newName);
                }
                catch (StorageException homeException)
                {
                    // even this path did not work
                    recoveryError = homeException;
                }
            }
        }

        if (recoveryError != null)
        {
            await StoreForRecovery(acctok.WopiUser, newName, context.TokenForLog(), recoveryError.Message, null, context);
            // conflict file was stored on recovery space, tell user (but reason is advisory...)
            reason += ", please contact support to recover it";
        }
        else
        {
            // otherwise, conflict file was saved to user space
            reason += ", conflict copy created";
        }

        // use a CONFLICT response as it is better handled by the app to signal the issue to the user
        return MakeConflictResponse(context, "PUTFILE", acctok.UserId, retrievedLock, requestedLock, "NA", acctok.FileName, reason);
    }

    /// <summary>
    /// The configured homepath of the user, cf. `homepath` in wopiserver.conf.
    /// </summary>
    public string HomePathFor(AccessToken acctok) =>
        state.HomePath
            .Replace("user_initial", acctok.WopiUser[..1])
            .Replace("username", acctok.WopiUserName);

    /// <summary>
    /// Stores the given content, or else the body of the request, in the local recovery space.
    /// </summary>
    public async Task StoreForRecovery(string wopiUser, string filename, string tokenForLog, string error, byte[]? content, HttpContext? context)
    {
        if (content == null)
        {
            if (context == null)
            {
                log.LogWarning(
                    "Requested to store for recovery but no content and no request context {filename} {token} {originalerror}",
                    filename,
                    tokenForLog,
                    error);
                return;
            }

            content = await ReadBody(context.Request);
        }

        var filePath = Path.Combine(
            state.RecoveryPath,
            $"{DateTime.Now.ToString("yyyyMMddTHHmmss", CultureInfo.InvariantCulture)}_editedby_{PyStrings.SecureFilename(wopiUser.Split('!')[0])}_origat_{PyStrings.SecureFilename(filename)}");
        try
        {
            await File.WriteAllBytesAsync(filePath, content);
            log.LogWarning(
                "Error writing file, a copy was stored locally for later recovery {filename} {recoveredpath} {token} {error}",
                filename,
                filePath,
                tokenForLog,
                error);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            log.LogCritical(
                "Error writing file and failed to recover it to local storage, data is LOST {filename} {token} {originalerror} {recoveryerror}",
                filename,
                tokenForLog,
                error,
                exception.Message);
        }
    }

    static async Task<byte[]> ReadBody(HttpRequest request)
    {
        if (request.Body.CanSeek)
        {
            request.Body.Position = 0;
        }

        using var memory = new MemoryStream();
        await request.Body.CopyToAsync(memory);
        return memory.ToArray();
    }
}

/// <param name="Inode">The file's inode, which serves as fileid.</param>
/// <param name="ViewMode">The view mode granted, possibly downgraded from the requested one.</param>
/// <param name="Details">The reason of a downgraded view mode.</param>
record GeneratedToken(string Inode, string AccessToken, ViewMode ViewMode, string Details);
