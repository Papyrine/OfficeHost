/// <summary>
/// Implementation of the core WOPI API.
/// </summary>
class WopiCore(WopiState state, WopiUtils utils, ILogger log)
{
    const string ioError = "I/O Error, please contact support";

    IStorage Storage => state.Storage;

    WopiConfig Config => state.Config;

    static string AccessTokenArg(HttpContext context) =>
        context.Request.Arg("access_token") ?? "";

    static string Separator(string url)
    {
        if (url.Contains('?'))
        {
            return "&";
        }

        return "?";
    }

    /// <summary>
    /// Implements the CheckFileInfo WOPI call.
    /// </summary>
    public async Task<WopiResult> CheckFileInfo(HttpContext context, string fileId, AccessToken acctok)
    {
        var token = context.TokenForLog();
        StatInfo statInfo;
        try
        {
            statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
        }
        catch (StorageException exception)
        {
            log.LogInformation("Requested file not found {filename} {token} {details}", acctok.FileName, token, exception.Message);
            return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
        }

        var viewMode = acctok.ViewMode;
        var isProxied = acctok.AppName == state.ProxiedAppName;

        // ** Populate general metadata for this file **
        var fmd = new JsonObject();
        var baseFileName = PosixPath.BaseName(acctok.FileName);
        fmd["BaseFileName"] = baseFileName;
        fmd["BreadcrumbDocName"] = baseFileName;
        // [MS-WOPI] 3.3.5.1.2.2: the value MUST begin with a ".", so a file without extension has none
        var extension = PosixPath.Extension(acctok.FileName);
        if (extension.Length > 0)
        {
            fmd["FileExtension"] = extension;
        }
        var wopiSrc = $"WOPISrc={utils.GenerateWopiSrc(fileId, isProxied)}&access_token={AccessTokenArg(context)}";
        var hostEditUrl = Config.Get("general", "hostediturl");
        if (!string.IsNullOrEmpty(hostEditUrl))
        {
            var editUrl = WopiUtils.GenerateUrlFromTemplate(hostEditUrl, acctok);
            fmd["HostEditUrl"] = editUrl;
            if (Uri.TryCreate(editUrl, UriKind.Absolute, out var uri))
            {
                fmd["PostMessageOrigin"] = $"{uri.Scheme}://{uri.Authority}";
            }

            fmd["EditModePostMessage"] = true;
            fmd["EditNotificationPostMessage"] = true;
        }
        else
        {
            fmd["HostEditUrl"] = $"{acctok.AppEditUrl}{Separator(acctok.AppEditUrl ?? "")}{wopiSrc}";
        }

        var hostViewUrl = Config.Get("general", "hostviewurl");
        if (!string.IsNullOrEmpty(hostViewUrl))
        {
            fmd["HostViewUrl"] = WopiUtils.GenerateUrlFromTemplate(hostViewUrl, acctok);
        }
        else
        {
            fmd["HostViewUrl"] = $"{acctok.AppViewUrl}{Separator(acctok.AppViewUrl ?? "")}{wopiSrc}";
        }

        var fileSharingUrl = Config.Get("general", "filesharingurl");
        if (!string.IsNullOrEmpty(fileSharingUrl))
        {
            fmd["FileSharingUrl"] = WopiUtils.GenerateUrlFromTemplate(fileSharingUrl, acctok);
            fmd["FileSharingPostMessage"] = true;
        }

        if (acctok.FolderUrl != "/")
        {
            var closeUrl = $"{acctok.FolderUrl}?scrollTo={baseFileName}";
            fmd["CloseUrl"] = closeUrl;
            fmd["BreadcrumbFolderUrl"] = closeUrl;
        }

        if (acctok.UserName.Length == 0 ||
            acctok.UserType == UserType.Anonymous)
        {
            fmd["IsAnonymousUser"] = true;
            fmd["UserFriendlyName"] = $"Guest {WopiUtils.RandomString(3)}";
            fmd["BreadcrumbFolderName"] = "Public share";
        }
        else
        {
            fmd["IsAnonymousUser"] = false;
            fmd["UserFriendlyName"] = acctok.UserName;
            fmd["BreadcrumbFolderName"] = acctok.UserType == UserType.Ocm ? "Federated share" : "Parent folder";
        }

        var downloadUrl = Config.Get("general", "downloadurl");
        if (!viewMode.IsViewOnly() &&
            !string.IsNullOrEmpty(downloadUrl))
        {
            var url = $"{downloadUrl}?access_token={AccessTokenArg(context)}";
            fmd["DownloadUrl"] = url;
            fmd["FileUrl"] = url;
        }

        // these are strings in [MS-WOPI] 3.3.5.1.2.2: unlike the Python implementation, do not send them as null
        var brandingName = Config.Get("general", "brandingname");
        if (brandingName != null)
        {
            fmd["BreadcrumbBrandName"] = brandingName;
        }

        var brandingUrl = Config.Get("general", "brandingurl");
        if (brandingUrl != null)
        {
            fmd["BreadcrumbBrandUrl"] = brandingUrl;
        }
        fmd["OwnerId"] = statInfo.OwnerId;
        // typically same as OwnerId; different when accessing shared documents
        fmd["UserId"] = acctok.WopiUserId;
        fmd["Size"] = statInfo.Size;
        fmd["LastModifiedTime"] = $"{DateTimeOffset.FromUnixTimeSeconds(statInfo.MTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}.000";
        // note that in ownCloud 10 the version is generated as: `'V' + etag + checksum`. Here etag is `version:checksum`.
        fmd["Version"] = $"v{statInfo.Etag}";
        fmd["AccessTokenExpiry"] = acctok.Exp * 1000;
        fmd["ServerTime"] = state.Now * 1000;

        fmd["SupportsExtendedLockLength"] = true;
        fmd["SupportsGetLock"] = true;
        fmd["SupportsCoauth"] = true;
        var canWrite = viewMode.CanWrite();
        fmd["SupportsUpdate"] = canWrite;
        fmd["UserCanWrite"] = canWrite;
        fmd["SupportsLocks"] = canWrite;
        fmd["SupportsDeleteFile"] = canWrite;
        fmd["ReadOnly"] = !canWrite;
        fmd["RestrictedWebViewOnly"] = viewMode.IsViewOnly();
        // SaveAs functionality is disabled for anonymous and federated users, as they have no personal space where to save
        // as an alternate location and we cannot assume that saving to the same folder is allowed (e.g. single-file shares).
        // Instead, regular (authenticated) users are offered a SaveAs (unless in view-only mode), where the operation
        // is executed to the user's home if no access is given to the same folder where the file is.
        fmd["UserCanNotWriteRelative"] = viewMode.IsViewOnly() || acctok.UserType != UserType.Regular;
        var canRename = state.EnableRename && canWrite;
        fmd["SupportsRename"] = canRename;
        fmd["UserCanRename"] = canRename;
        fmd["SupportsUserInfo"] = true;
        if (statInfo.Xattrs.TryGetValue($"{WopiUtils.UserInfoKey}.{acctok.WopiUserName}", out var userInfo) &&
            userInfo.Length > 0)
        {
            fmd["UserInfo"] = userInfo;
        }

        // ** Populate app-specific metadata **
        if (Config.IsTrue("apps", "businessflow", true))
        {
            // according to Microsoft, this must be enabled for all users
            fmd["LicenseCheckForEditIsEnabled"] = true;
        }

        if (Config.IsTrue("apps", "earlyfeatures", false))
        {
            fmd["AllowEarlyFeatures"] = true;
        }

        var complianceDomain = Config.Get("apps", "compliancedomain");
        if (complianceDomain != null)
        {
            fmd["ComplianceDomainPrefix"] = complianceDomain;
        }

        // 28/11/2025: the following properties are CSPP+ related and necessary only when running the wopivalidator,
        // with 'SequenceNumber' even breaking normal CSPP Office apps:
        // learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/rest/files/checkfileinfo/checkfileinfo-csppp
        var isWopiTest = baseFileName.Contains(".wopitest");
        if (isWopiTest)
        {
            fmd["SharingStatus"] = "Shared";
            fmd["FileGeoLocationCode"] = complianceDomain ?? "euc";
            fmd["RealTimeChannelEndpointUrl"] = "";
            fmd["OfficeCollaborationServiceEndpointUrl"] = "";
            var colon = statInfo.Etag.IndexOf(':');
            if (colon >= 0 &&
                long.TryParse(statInfo.Etag.AsSpan(0, colon), CultureInfo.InvariantCulture, out var sequenceNumber))
            {
                fmd["SequenceNumber"] = sequenceNumber;
            }
            else
            {
                log.LogWarning("Failed to extract SequenceNumber from etag {token} {etag}", token, statInfo.Etag);
            }
        }

        // the following is to enable the 'Edit in Word/Excel/PowerPoint' (desktop) action (probably broken);
        // if no WebDAV URL is provided, this setting is ignored
        var webDavUrl = Config.Get("general", "webdavurl");
        if (webDavUrl != null)
        {
            fmd["ClientUrl"] = $"{webDavUrl}/{acctok.FileName}";
        }

        // postMessage extensions for non-Microsoft editors: we expose them in all cases, except when running the wopivalidator test suite
        if (!isWopiTest)
        {
            fmd["EnableOwnerTermination"] = true;
            fmd["IsAdminUser"] = false;
            fmd["EnableInsertRemoteImage"] = true;
            fmd["EnableInsertRemoteFile"] = true;
            fmd["EnableRemoteLinkPicker"] = true;
            // @mentions in comments, and the "copy link to this comment" action that goes with them
            fmd["UserMentionPostMessage"] = true;
            var viewOnly = viewMode.IsViewOnly();
            fmd["DisableExport"] = viewOnly;
            fmd["DisableCopy"] = viewOnly;
            fmd["DisablePrint"] = viewOnly;
            fmd["HideExportOption"] = viewOnly;
        }

        // ** Response and log **
        var response = WopiResult.Json(fmd);
        if (!Config.IsDebug)
        {
            foreach (var key in new[] { "HostViewUrl", "HostEditUrl", "DownloadUrl", "FileUrl", "BreadcrumbBrandUrl", "FileSharingUrl" })
            {
                fmd[key] = "_redacted_";
            }
        }

        log.LogInformation("File metadata response {token} {metadata}", token, fmd.ToJsonString(Json.Options));
        return response;
    }

    /// <summary>
    /// Implements the GetFile WOPI call.
    /// </summary>
    public async Task<WopiResult> GetFile(HttpContext context, AccessToken acctok)
    {
        try
        {
            // TODO for the time being we do not look if the file is locked. Once exclusive locks are implemented in Reva,
            // the lock must be fetched prior to the following call in order to access the file.
            // stat the file to get the current version
            var statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
            // stream file from storage to client
            var stream = await Storage.ReadFile(acctok.Endpoint, acctok.FileName, acctok.UserId, null);
            var response = new WopiResult(StatusCodes.Status200OK)
            {
                ContentType = "application/octet-stream",
                BodyStream = stream
            };
            response.Headers["Content-Disposition"] = $"attachment; filename*=UTF-8''{Uri.EscapeDataString(PosixPath.BaseName(acctok.FileName))}";
            response.Headers["X-Frame-Options"] = "sameorigin";
            response.Headers["X-XSS-Protection"] = "1; mode=block";
            response.Headers["X-WOPI-ItemVersion"] = $"v{statInfo.Etag}";
            return response;
        }
        catch (StorageException exception)
        {
            // File got deleted meanwhile, or some other remote I/O error
            log.LogError(
                "GetFile: download failed {endpoint} {filename} {token} {error}",
                acctok.Endpoint,
                acctok.FileName,
                context.TokenForLog(),
                exception.Message);
            // [MS-WOPI] 3.3.5.3.1: 404 is the status for an unknown file
            if (exception.IsNotFound)
            {
                return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
            }

            return WopiResult.Message("Failed to access file from storage", StatusCodes.Status500InternalServerError);
        }
    }

    // The following operations are all called on POST /wopi/files/<fileid>

    /// <summary>
    /// Implements the Lock, RefreshLock, and UnlockAndRelock WOPI calls.
    /// </summary>
    public async Task<WopiResult> SetLock(HttpContext context, string fileId, AccessToken acctok)
    {
        var request = context.Request;
        var token = context.TokenForLog();
        var operation = request.Header("X-WOPI-Override")!;
        var requestedLock = request.Header("X-WOPI-Lock") ?? "";
        var oldLock = request.Header("X-WOPI-OldLock");
        var validateTarget = request.Header("X-WOPI-Validate-Target");
        var filename = acctok.FileName;
        var (retrievedLock, lockHolder) = await utils.RetrieveWopiLock(fileId, operation, requestedLock, acctok.Endpoint, filename, acctok.UserId, token);
        string? saveTime = null;

        StatInfo statInfo;
        try
        {
            // validate that the underlying file is still there (it might have been moved/deleted)
            statInfo = await Storage.StatX(acctok.Endpoint, filename, acctok.UserId);
        }
        catch (StorageException exception)
        {
            log.LogWarning("Error with target file {lockop} {filename} {token} {error}", operation.Title(), filename, token, exception.Message);
            if (exception.IsNotFound)
            {
                return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
            }

            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }

        if (!string.IsNullOrEmpty(retrievedLock) ||
            operation == "REFRESH_LOCK")
        {
            // useful for later checks
            saveTime = statInfo.Xattrs.GetValueOrDefault(WopiUtils.LastSaveTimeKey);
            if (!IsValidSaveTime(saveTime, statInfo.MTime))
            {
                // we had stale information, discard
                log.LogWarning("Detected external modification {filename} {savetime} {mtime} {token}", filename, saveTime, statInfo.MTime, token);
                saveTime = null;
            }
        }

        // perform the required checks for the validity of the new lock
        if (operation == "REFRESH_LOCK" &&
            string.IsNullOrEmpty(retrievedLock) &&
            (string.IsNullOrEmpty(validateTarget) || saveTime == null))
        {
            // validateTarget is an extension of the API: a REFRESH_LOCK without previous lock but with a Validate-Target header
            // is allowed, provided that the target file was last saved by WOPI (i.e. savetime is valid) and not overwritten
            // by other external actions (cf. PutFile logic)
            return utils.MakeConflictResponse(
                context,
                operation,
                acctok.UserId,
                null,
                requestedLock,
                oldLock,
                filename,
                $"The file was not locked{(string.IsNullOrEmpty(validateTarget) ? "" : " and got modified")}",
                saveTime);
        }

        if (!string.IsNullOrEmpty(oldLock) &&
            string.IsNullOrEmpty(retrievedLock))
        {
            // UnlockAndRelock releases the old lock, so it fails when there is none ([MS-WOPI] 3.3.5.1.23)
            return utils.MakeConflictResponse(context, operation, acctok.UserId, null, requestedLock, oldLock, filename, "The file was not locked", saveTime);
        }

        // now check for and create an "external" lock if required
        if (Config.IsTrue("general", "detectexternallocks", true) &&
            state.CodeTypes.Contains(PosixPath.Extension(filename)))
        {
            if (retrievedLock == WopiUtils.ExternalLock)
            {
                return utils.MakeConflictResponse(context, operation, acctok.UserId, retrievedLock, requestedLock, oldLock, filename, $"The file is locked by {lockHolder}", saveTime);
            }

            var conflict = await CreateLibreOfficeLock(context, operation, acctok, requestedLock, oldLock, saveTime);
            if (conflict != null)
            {
                return conflict;
            }
        }

        try
        {
            // LOCK or REFRESH_LOCK: atomically set the lock to the given one, including the expiration time,
            // and return conflict response if the file was already locked
            await Storage.SetLock(acctok.Endpoint, filename, acctok.UserId, acctok.AppName, WopiUtils.EncodeLock(requestedLock));

            // on first lock, set in addition an xattr with the current time for later conflicts checking if required
            try
            {
                await Storage.SetXattr(
                    acctok.Endpoint,
                    filename,
                    acctok.UserId,
                    WopiUtils.LastSaveTimeKey,
                    state.Now.ToString(CultureInfo.InvariantCulture),
                    new(acctok.AppName, WopiUtils.EncodeLock(requestedLock)));
            }
            catch (StorageException exception)
            {
                // not fatal, but will generate a conflict file later on, so log a warning
                log.LogWarning(
                    "Unable to set lastwritetime xattr {lockop} {user} {filename} {token} {reason}",
                    operation.Title(),
                    acctok.UserId.Short(),
                    filename,
                    token,
                    exception.Message);
            }

            return utils.MakeLockSuccessResponse(context, operation, acctok, requestedLock, oldLock, $"v{statInfo.Etag}");
        }
        catch (StorageException exception)
        {
            if (exception.IsLockMismatch)
            {
                // another session was faster than us, or the file was already WOPI-locked:
                // get the lock that was set
                if (string.IsNullOrEmpty(retrievedLock))
                {
                    (retrievedLock, lockHolder) = await utils.RetrieveWopiLock(fileId, operation, requestedLock, acctok.Endpoint, filename, acctok.UserId, token);
                }

                // validate against either the given lock (RefreshLock case) or the given old lock (UnlockAndRelock case);
                // in the context of the EXCL_ERROR case, retrievedLock may be None only if the storage is holding a user lock
                if (string.IsNullOrEmpty(retrievedLock) ||
                    !utils.CompareWopiLocks(retrievedLock, string.IsNullOrEmpty(oldLock) ? requestedLock : oldLock))
                {
                    // lock mismatch, the WOPI client is supposed to acknowledge the existing lock to start a collab session,
                    // or deny access to the file in edit mode otherwise
                    return utils.MakeConflictResponse(
                        context,
                        operation,
                        acctok.UserId,
                        retrievedLock,
                        requestedLock,
                        oldLock,
                        filename,
                        $"The file is locked by {(string.IsNullOrEmpty(lockHolder) ? "another editor" : lockHolder)}",
                        saveTime);
                }

                // else it's our own lock, refresh it (rechecking the oldLock if necessary, for atomicity) and return
                try
                {
                    await Storage.RefreshLock(acctok.Endpoint, filename, acctok.UserId, acctok.AppName, WopiUtils.EncodeLock(requestedLock), WopiUtils.EncodeLock(oldLock));
                    return utils.MakeLockSuccessResponse(context, operation, acctok, requestedLock, oldLock, $"v{statInfo.Etag}");
                }
                catch (StorageException refreshException)
                {
                    // this is unexpected now
                    log.LogError(
                        "Failed to refresh lock {lockop} {filename} {token} {lock} {error}",
                        operation.Title(),
                        filename,
                        token,
                        requestedLock,
                        refreshException.Message);
                }
            }

            // any other error is raised
            log.LogError("Unable to store WOPI lock {lockop} {filename} {token} {lock} {error}", operation.Title(), filename, token, requestedLock, exception.Message);
            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }
    }

    static bool IsValidSaveTime(string? saveTime, long mtime) =>
        saveTime != null &&
        saveTime.Length > 0 &&
        saveTime.All(char.IsAsciiDigit) &&
        long.TryParse(saveTime, CultureInfo.InvariantCulture, out var saved) &&
        saved >= mtime;

    /// <summary>
    /// Creates a LibreOffice-compatible lock file for interoperability purposes, making sure to
    /// not overwrite any existing or being created lock. Returns a conflict response if a valid
    /// LibreOffice lock is held by someone else.
    /// </summary>
    async Task<WopiResult?> CreateLibreOfficeLock(HttpContext context, string operation, AccessToken acctok, string requestedLock, string? oldLock, string? saveTime)
    {
        var filename = acctok.FileName;
        var lockName = WopiUtils.GetLibreOfficeLockName(filename);
        try
        {
            var lockContent = string.Create(
                CultureInfo.InvariantCulture,
                $",{acctok.WopiUserName} via web collaboration,{Config.Get("general", "brandingname", "cernbox").ToLowerInvariant()},{DateTime.Now:dd.MM.yyyy HH:mm},wopiserver;");
            using var content = new MemoryStream(Encoding.UTF8.GetBytes(lockContent));
            await Storage.WriteFile(acctok.Endpoint, lockName, acctok.UserId, content, content.Length, null, isLock: true);
            return null;
        }
        catch (StorageException exception) when (exception.IsLockMismatch)
        {
            // retrieve the LibreOffice-compatible lock just found
            string retrievedLoLock;
            try
            {
                retrievedLoLock = await utils.ReadText(acctok.Endpoint, lockName, acctok.UserId);
                // check that the lock is not stale
                if (LibreOfficeLockTime(retrievedLoLock) + Config.LockExpiration < state.Now)
                {
                    retrievedLoLock = "WOPIServer";
                }
            }
            catch (Exception readException) when (readException is StorageException or FormatException or IndexOutOfRangeException)
            {
                // could not read the lock, assume it expired and take ownership
                retrievedLoLock = "WOPIServer";
            }

            if (retrievedLoLock.Contains("WOPIServer"))
            {
                // it's our previous lock or it had expired: all right, move on
                return null;
            }

            // the file was externally locked, make this call fail
            var lockHolder = retrievedLoLock.Contains(',') ? retrievedLoLock.Split(',')[1] : "";
            log.LogWarning(
                "Valid LibreOffice lock found, denying WOPI lock {lockop} {filename} {holder}",
                operation.Title(),
                filename,
                lockHolder.Length > 0 ? lockHolder : retrievedLoLock);
            var reason = $"File locked by {(lockHolder.Length > 0 ? $"{lockHolder} via LibreOffice" : "a LibreOffice user")}";
            return utils.MakeConflictResponse(context, operation, acctok.UserId, "External App", requestedLock, oldLock, filename, reason, saveTime);
        }
        catch (StorageException exception)
        {
            // any other error is logged but not raised as this is optimistically not blocking WOPI operations
            // this includes the case of access denied (over)writing the LibreOffice lock because of accessing
            // a single-file share
            log.LogWarning(
                "Failed to store LibreOffice-compatible lock {lockop} {filename} {token} {error}",
                operation.Title(),
                filename,
                context.TokenForLog(),
                exception.Message);
            return null;
        }
    }

    /// <summary>
    /// The time, in Unix seconds, of a LibreOffice lock like `,user,host,18.10.2026 15:30,LibreOffice;`.
    /// </summary>
    static long LibreOfficeLockTime(string loLock) =>
        new DateTimeOffset(DateTime.ParseExact(loLock.Split(',')[3], "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture)).ToUnixTimeSeconds();

    /// <summary>
    /// Implements the GetLock WOPI call.
    /// </summary>
    public async Task<WopiResult> GetLock(HttpContext context, string fileId, AccessToken acctok)
    {
        if (!await Exists(acctok, acctok.FileName))
        {
            return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
        }

        var (retrievedLock, _) = await utils.RetrieveWopiLock(fileId, "GETLOCK", "", acctok.Endpoint, acctok.FileName, acctok.UserId, context.TokenForLog());
        // an unlocked file is reported with an empty lock, unlike the Python implementation which returns 404:
        // [MS-WOPI] 3.3.5.1.10 reserves 404 for an unknown file, and the WOPI validator expects 200
        var response = new WopiResult(StatusCodes.Status200OK);
        response.Headers["X-WOPI-Lock"] = retrievedLock ?? "";
        return response;
    }

    /// <summary>
    /// Implements the Unlock WOPI call.
    /// </summary>
    public async Task<WopiResult> Unlock(HttpContext context, string fileId, AccessToken acctok)
    {
        var requestedLock = context.Request.Header("X-WOPI-Lock") ?? "";
        if (!await Exists(acctok, acctok.FileName))
        {
            // [MS-WOPI] 3.3.5.1.22: an unknown file is a 404 rather than a lock mismatch
            return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
        }

        var (retrievedLock, _) = await utils.RetrieveWopiLock(fileId, "UNLOCK", requestedLock, acctok.Endpoint, acctok.FileName, acctok.UserId, context.TokenForLog());
        if (!utils.CompareWopiLocks(retrievedLock, requestedLock))
        {
            return utils.MakeConflictResponse(context, "UNLOCK", acctok.UserId, retrievedLock, requestedLock, "NA", acctok.FileName, "Lock mismatch unlocking file");
        }

        // OK, the lock matches, remove it
        StatInfo statInfo;
        try
        {
            // validate that the underlying file is still there
            statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
            await Storage.Unlock(acctok.Endpoint, acctok.FileName, acctok.UserId, acctok.AppName, WopiUtils.EncodeLock(requestedLock));
            // note we leave the lastwritetime xattr, as it may be useful for RefreshLock with X-WOPI-ValidateTarget headers
        }
        catch (StorageException exception)
        {
            if (exception.IsNotFound)
            {
                return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
            }

            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }

        if (Config.IsTrue("general", "detectexternallocks", true))
        {
            try
            {
                // also remove the LibreOffice-compatible lock file when relevant
                if (!state.NonOfficeTypes.Contains(PosixPath.Extension(acctok.FileName)))
                {
                    await Storage.RemoveFile(acctok.Endpoint, WopiUtils.GetLibreOfficeLockName(acctok.FileName), acctok.UserId, force: true);
                }
            }
            catch (StorageException)
            {
                // ignore, it's not worth to report anything here
            }
        }

        // and update our internal lists of opened files and conflicted sessions
        state.CloseFile(acctok.FileName, context.Request.Header("X-WOPI-SessionId"));
        var response = new WopiResult(StatusCodes.Status200OK);
        response.Headers["X-WOPI-ItemVersion"] = $"v{statInfo.Etag}";
        return response;
    }

    /// <summary>
    /// Implements the PutRelative WOPI call. Corresponds to the 'Save as...' menu entry.
    /// </summary>
    public async Task<WopiResult> PutRelative(HttpContext context, string fileId, AccessToken acctok)
    {
        var request = context.Request;
        var token = context.TokenForLog();
        var suggestedTarget = request.Header("X-WOPI-SuggestedTarget");
        var relativeTarget = request.Header("X-WOPI-RelativeTarget");
        var overwriteTarget = string.Equals(request.Header("X-WOPI-OverwriteRelativeTarget"), "TRUE", StringComparison.OrdinalIgnoreCase);
        log.LogInformation(
            "PutRelative {user} {filename} {fileid} {suggTarget} {relTarget} {overwrite} {wopitimestamp} {token}",
            acctok.UserId.Short(),
            acctok.FileName,
            fileId,
            suggestedTarget,
            relativeTarget,
            overwriteTarget,
            request.Header("X-WOPI-TimeStamp"),
            token);

        if (acctok.ViewMode.IsViewOnly() ||
            acctok.UserType != UserType.Regular)
        {
            // UNAUTHORIZED may seem better but the WOPI validator tests explicitly expect NOT_IMPLEMENTED
            return WopiResult.Message("Unauthorized to perform PutRelative", StatusCodes.Status501NotImplemented);
        }

        // either one xor the other MUST be present
        if (string.IsNullOrEmpty(suggestedTarget) == string.IsNullOrEmpty(relativeTarget))
        {
            return WopiResult.Message("Conflicting headers given", StatusCodes.Status400BadRequest);
        }

        var folder = PosixPath.DirName(acctok.FileName);
        string targetName;
        if (!string.IsNullOrEmpty(suggestedTarget))
        {
            // the suggested target is a UTF7-encoded (!) filename that can be changed to avoid collisions
            // and, as [MS-WOPI] 3.3.5.1.15 requires, to be legally named
            suggestedTarget = LegalName(PyStrings.DecodeUtf7(suggestedTarget));
            if (suggestedTarget[0] == '.')
            {
                // we just have the extension here
                targetName = PosixPath.Join(folder, $"{PosixPath.BaseName(PosixPath.SplitExt(acctok.FileName).Root)}{suggestedTarget}");
            }
            else
            {
                targetName = PosixPath.Join(folder, suggestedTarget);
            }

            // check for existence of the target file and adjust until a non-existing one is obtained
            var freeName = await FreeName(acctok, targetName, token);
            if (freeName == null)
            {
                return WopiResult.Message("Error with the given target", StatusCodes.Status500InternalServerError);
            }

            targetName = freeName;
        }
        else
        {
            // the relative target is a UTF7-encoded filename to be respected, and that may overwrite an existing file
            var relativeName = PyStrings.DecodeUtf7(relativeTarget!);
            if (!IsLegalName(relativeName))
            {
                // the name must not be modified, so an illegal one fails the request
                return WopiResult.Message("Specified name is illegal", StatusCodes.Status400BadRequest);
            }

            targetName = PosixPath.Join(folder, relativeName);
            try
            {
                // check for file existence
                var statInfo = await Storage.StatX(acctok.Endpoint, targetName, acctok.UserId);
                // the file exists, check lock
                var (retrievedTargetLock, _) = await utils.RetrieveWopiLock(fileId, "PUT_RELATIVE", null, acctok.Endpoint, targetName, acctok.UserId, token);
                // deny if lock is valid or if overwriteTarget is False
                if (!overwriteTarget ||
                    !string.IsNullOrEmpty(retrievedTargetLock))
                {
                    var conflict = new JsonObject
                    {
                        ["message"] = "Target file already exists",
                        // specs (the WOPI validator) require these to be populated with valid values
                        ["Name"] = PosixPath.BaseName(targetName),
                        ["Url"] = utils.GenerateWopiSrc(statInfo.Inode, acctok.AppName == state.ProxiedAppName)
                    };
                    var conflictResponse = utils.MakeConflictResponse(context, "PUT_RELATIVE", acctok.UserId, retrievedTargetLock, "NA", "NA", targetName, conflict);
                    // a name the client can use in a further PutRelativeFile
                    var validTarget = await FreeName(acctok, targetName, token);
                    if (validTarget != null)
                    {
                        conflictResponse.Headers["X-WOPI-ValidRelativeTarget"] = PyStrings.EncodeUtf7(PosixPath.BaseName(validTarget));
                    }

                    return conflictResponse;
                }
            }
            catch (StorageException)
            {
                // optimistically assume we're clear
            }
        }

        // either way, we now have a targetName to save the file: attempt to do so
        try
        {
            await utils.StoreWopiFile(context, acctok, null, WopiUtils.LastSaveTimeKey, targetName);
        }
        catch (StorageException exception)
        {
            if (!exception.IsAccessDenied)
            {
                return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
            }

            var raiseNoAccess = true;
            // make an attempt in the user's home if possible: that would be allowed for regular (authenticated) users
            // when the target is a single file r/w share
            if (acctok.UserType == UserType.Regular)
            {
                targetName = PosixPath.Join(utils.HomePathFor(acctok), PosixPath.BaseName(targetName));
                log.LogInformation("PutRelative: set homepath as destination {user} {filename} {target} {token}", acctok.UserId.Short(), acctok.FileName, targetName, token);
                try
                {
                    await utils.StoreWopiFile(context, acctok, null, WopiUtils.LastSaveTimeKey, targetName);
                    raiseNoAccess = false;
                }
                catch (StorageException)
                {
                    // at this point give up and return error
                }
            }

            if (raiseNoAccess)
            {
                return WopiResult.Message("Unauthorized to perform PutRelative", StatusCodes.Status501NotImplemented);
            }
        }

        // generate an access token for the new file
        log.LogInformation(
            "PutRelative: generating new access token {user} {filename} {mode} {friendlyname}",
            acctok.UserId.Short(),
            targetName,
            "ViewMode.READ_WRITE",
            acctok.UserName);
        var generated = await utils.GenerateAccessToken(
            acctok.UserId,
            targetName,
            ViewMode.ReadWrite,
            acctok.UserName,
            acctok.WopiUser,
            acctok.UserType,
            acctok.FolderUrl,
            acctok.Endpoint,
            acctok.AppName,
            acctok.AppEditUrl,
            acctok.AppViewUrl,
            acctok.Trace ?? "N/A");
        // prepare and send the response as JSON
        var (_, newFileId) = Common.DecodeInode(generated.Inode);
        var newWopiSrc = $"{utils.GenerateWopiSrc(generated.Inode, acctok.AppName == state.ProxiedAppName)}&access_token={generated.AccessToken}";
        var putRelativeMetadata = new JsonObject
        {
            ["Name"] = PosixPath.BaseName(targetName),
            ["Url"] = UrlCodec.UnquotePlus(newWopiSrc).Replace("&access_token", "?access_token")
        };
        var hostEditUrl = Config.Get("general", "hostediturl");
        if (!string.IsNullOrEmpty(hostEditUrl))
        {
            putRelativeMetadata["HostEditUrl"] = WopiUtils.GenerateUrlFromTemplate(hostEditUrl, acctok.UserName, acctok.UserType, acctok.FolderUrl, targetName, acctok.Endpoint, newFileId, acctok.AppName);
        }
        else
        {
            putRelativeMetadata["HostEditUrl"] = $"{acctok.AppEditUrl}{Separator(acctok.AppEditUrl ?? "")}WOPISrc={newWopiSrc}";
        }

        var hostViewUrl = Config.Get("general", "hostviewurl");
        if (!string.IsNullOrEmpty(hostViewUrl))
        {
            putRelativeMetadata["HostViewUrl"] = WopiUtils.GenerateUrlFromTemplate(hostViewUrl, acctok.UserName, acctok.UserType, acctok.FolderUrl, targetName, acctok.Endpoint, newFileId, acctok.AppName);
        }
        else
        {
            putRelativeMetadata["HostViewUrl"] = $"{acctok.AppViewUrl}{Separator(acctok.AppViewUrl ?? "")}WOPISrc={newWopiSrc}";
        }

        var response = WopiResult.Json(putRelativeMetadata);
        putRelativeMetadata["Url"] = "_redacted_";
        putRelativeMetadata["HostEditUrl"] = "_redacted_";
        putRelativeMetadata["HostViewUrl"] = "_redacted_";
        log.LogInformation("PutRelative response {token} {metadata}", generated.AccessToken.Short(), putRelativeMetadata.ToJsonString(Json.Options));
        return response;
    }

    /// <summary>
    /// The given path, or the first variant with `_copy` suffixes that does not exist. Null on any other storage error.
    /// </summary>
    async Task<string?> FreeName(AccessToken acctok, string targetName, string token)
    {
        while (true)
        {
            try
            {
                await Storage.Stat(acctok.Endpoint, targetName, acctok.UserId);
                // the file exists: try a different name
                var (name, extension) = PosixPath.SplitExt(targetName);
                targetName = $"{name}_copy{extension}";
            }
            catch (StorageException exception)
            {
                if (exception.IsNotFound)
                {
                    // OK, the targetName is good to go
                    return targetName;
                }

                // we got another error with this file, fail
                log.LogError("Error in PutRelative {user} {filename} {token} {error}", acctok.UserId.Short(), targetName, token, exception.Message);
                return null;
            }
        }
    }

    /// <summary>
    /// Whether a file name given by a WOPI client designates a file in the current folder.
    /// </summary>
    static bool IsLegalName(string name) =>
        name.Length > 0 &&
        name is not ("." or "..") &&
        name.IndexOfAny(['/', '\\']) < 0 &&
        !name.Any(char.IsControl);

    /// <summary>
    /// The name with the characters that would make it illegal replaced.
    /// </summary>
    static string LegalName(string name)
    {
        var legal = new string(name.Select(_ => _ is '/' or '\\' || char.IsControl(_) ? '_' : _).ToArray());
        if (legal is "" or "." or "..")
        {
            return "_";
        }

        return legal;
    }

    /// <summary>
    /// Implements the DeleteFile WOPI call.
    /// </summary>
    public async Task<WopiResult> DeleteFile(HttpContext context, string fileId, AccessToken acctok)
    {
        var (retrievedLock, _) = await utils.RetrieveWopiLock(fileId, "DELETE", "", acctok.Endpoint, acctok.FileName, acctok.UserId, context.TokenForLog());
        if (retrievedLock != null)
        {
            // file is locked and cannot be deleted
            return utils.MakeConflictResponse(context, "DELETE", acctok.UserId, retrievedLock, "NA", "NA", acctok.FileName, "Cannot delete a locked file");
        }

        try
        {
            await Storage.RemoveFile(acctok.Endpoint, acctok.FileName, acctok.UserId);
            return WopiResult.Message("OK", StatusCodes.Status200OK);
        }
        catch (StorageException exception)
        {
            if (exception.IsNotFound)
            {
                return WopiResult.Message("File not found", StatusCodes.Status404NotFound);
            }

            log.LogError("DeleteFile {token} {error}", context.TokenForLog(), exception.Message);
            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Implements the RenameFile WOPI call.
    /// </summary>
    public async Task<WopiResult> RenameFile(HttpContext context, string fileId, AccessToken acctok)
    {
        var request = context.Request;
        var requestedName = request.Header("X-WOPI-RequestedName");
        if (requestedName == null)
        {
            log.LogWarning("Missing argument {client} {requestedUrl} {error} {token}", context.RemoteAddress(), request.BaseUrl(), "'X-WOPI-RequestedName'", context.TokenForLog());
            return WopiResult.Message("Missing argument", StatusCodes.Status400BadRequest);
        }

        if (!state.EnableRename)
        {
            // CheckFileInfo does not advertise SupportsRename
            return WopiResult.Message("RenameFile is not enabled", StatusCodes.Status501NotImplemented);
        }

        var targetName = PyStrings.DecodeUtf7(requestedName);
        if (!IsLegalName(targetName))
        {
            return InvalidFileName("Specified name is illegal");
        }

        var renameMetadata = new JsonObject
        {
            ["Name"] = PosixPath.SplitExt(targetName).Root
        };
        // may not be specified
        var requestedLock = request.Header("X-WOPI-Lock");
        var (retrievedLock, _) = await utils.RetrieveWopiLock(fileId, "RENAMEFILE", requestedLock, acctok.Endpoint, acctok.FileName, acctok.UserId, context.TokenForLog());
        if (retrievedLock != null &&
            !utils.CompareWopiLocks(retrievedLock, requestedLock))
        {
            return utils.MakeConflictResponse(context, "RENAMEFILE", acctok.UserId, retrievedLock, requestedLock, "NA", acctok.FileName, "Lock mismatch renaming file");
        }

        try
        {
            // the destination name comes without base path and typically without extension
            targetName = PosixPath.Join(PosixPath.DirName(acctok.FileName), targetName);
            if (!PosixPath.BaseName(targetName).Contains('.'))
            {
                targetName += PosixPath.Extension(acctok.FileName);
            }

            log.LogInformation("RenameFile {user} {filename} {token} {targetname}", acctok.UserId.Short(), acctok.FileName, context.TokenForLog(), targetName);
            if (await Exists(acctok, targetName))
            {
                // a rename must not overwrite another file
                return InvalidFileName("A file with the requested name already exists");
            }

            // try to rename and pass the lock if present. Note that WOPI specs do not require files to be locked
            // on rename operations, but the backend may still fail as renames may be implemented as copy + delete,
            // which may require to pass a lock.
            LockMetadata? lockMetadata = null;
            if (!string.IsNullOrEmpty(retrievedLock))
            {
                lockMetadata = new(acctok.AppName, WopiUtils.EncodeLock(retrievedLock));
            }

            await Storage.RenameFile(acctok.Endpoint, acctok.FileName, targetName, acctok.UserId, lockMetadata);
            // also rename the LO lock if applicable
            if (state.CodeTypes.Contains(PosixPath.Extension(acctok.FileName)))
            {
                await Storage.RenameFile(
                    acctok.Endpoint,
                    WopiUtils.GetLibreOfficeLockName(acctok.FileName),
                    WopiUtils.GetLibreOfficeLockName(targetName),
                    acctok.UserId,
                    null);
            }

            // send the response as JSON
            return WopiResult.Json(renameMetadata);
        }
        catch (StorageException exception)
        {
            log.LogWarning("RenameFile {token} {error}", context.TokenForLog(), exception.Message);
            WopiResult response;
            if (exception.IsNotFound)
            {
                response = new(StatusCodes.Status404NotFound);
                response.Headers["X-WOPI-InvalidFileNameError"] = "File not found";
            }
            else if (exception.IsLockMismatch)
            {
                response = new(StatusCodes.Status501NotImplemented);
                response.Headers["X-WOPI-InvalidFileNameError"] = "Cannot rename/move unlocked file";
            }
            else
            {
                response = new(StatusCodes.Status500InternalServerError);
                response.Headers["X-WOPI-InvalidFileNameError"] = $"Failed to rename: {exception.Message}";
            }

            return response;
        }
    }

    static WopiResult InvalidFileName(string reason)
    {
        var response = new WopiResult(StatusCodes.Status400BadRequest);
        response.Headers["X-WOPI-InvalidFileNameError"] = reason;
        return response;
    }

    async Task<bool> Exists(AccessToken acctok, string filename)
    {
        try
        {
            await Storage.Stat(acctok.Endpoint, filename, acctok.UserId);
            return true;
        }
        catch (StorageException exception) when (exception.IsNotFound)
        {
            return false;
        }
    }

    /// <summary>
    /// Implements the editnew action as part of the PutFile WOPI call.
    /// </summary>
    async Task<WopiResult> CreateNewFile(HttpContext context, string fileId, AccessToken acctok)
    {
        var token = context.TokenForLog();
        log.LogInformation("PutFile {user} {filename} {fileid} {action} {token}", acctok.UserId.Short(), acctok.FileName, fileId, "editnew", token);
        try
        {
            // try to stat the file: a 0-size file is equivalent to not existing
            if ((await Storage.Stat(acctok.Endpoint, acctok.FileName, acctok.UserId)).Size != 0)
            {
                log.LogWarning("PutFile {error} {filename} {token}", "File exists but no WOPI lock provided", acctok.FileName, token);
                return WopiResult.Message("File exists", StatusCodes.Status409Conflict);
            }
        }
        catch (StorageException)
        {
        }

        // indeed the file did not exist, so we write it for the first time
        try
        {
            await utils.StoreWopiFile(context, acctok, null, WopiUtils.LastSaveTimeKey);
            log.LogInformation("File stored successfully {action} {user} {filename} {token}", "editnew", acctok.UserId.Short(), acctok.FileName, token);
            return WopiResult.Message("OK", StatusCodes.Status200OK);
        }
        catch (StorageException exception)
        {
            await utils.StoreForRecovery(acctok.WopiUser, acctok.FileName, token, exception.Message, null, context);
            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Implements the PutFile WOPI call.
    /// </summary>
    public async Task<WopiResult> PutFile(HttpContext context, string fileId, AccessToken acctok)
    {
        var requestedLock = context.Request.Header("X-WOPI-Lock");
        if (requestedLock == null)
        {
            // no lock given: assume we are in creation mode (cf. editnew WOPI action)
            return await CreateNewFile(context, fileId, acctok);
        }

        var token = context.TokenForLog();
        // otherwise, check that the caller holds the current lock on the file
        var (retrievedLock, lockHolder) = await utils.RetrieveWopiLock(fileId, "PUTFILE", requestedLock, acctok.Endpoint, acctok.FileName, acctok.UserId, token);
        if (retrievedLock == null)
        {
            return utils.MakeConflictResponse(context, "PUTFILE", acctok.UserId, retrievedLock, requestedLock, "NA", acctok.FileName, "Cannot overwrite unlocked file");
        }

        if (retrievedLock == WopiUtils.ExternalLock)
        {
            // this should not happen and we must fail, yet we save the file as conflict for the user to recover it
            log.LogError("Detected external lock, forcing conflict {user} {filename} {token}", acctok.UserId.Short(), acctok.FileName, token);
            return await utils.StoreAfterConflict(context, acctok, retrievedLock, requestedLock, $"Cannot overwrite file edited by {lockHolder}");
        }

        if (!utils.CompareWopiLocks(retrievedLock, requestedLock))
        {
            // the save operation is to be refused, but we should get a subsequent PutFile call with the correct lock, given that
            // the current lock is from WOPI; yet we keep the file in the recovery area in case the error turned out to be real
            await utils.StoreForRecovery(acctok.WopiUser, acctok.FileName, token, "Mismatched lock on PutFile", null, context);
            return utils.MakeConflictResponse(context, "PUTFILE", acctok.UserId, retrievedLock, requestedLock, "NA", acctok.FileName, $"Cannot overwrite file locked by {lockHolder}");
        }

        // OK, we can save the file: check the destination file against conflicts if required
        log.LogInformation("PutFile {user} {filename} {fileid} {action} {token}", acctok.UserId.Short(), acctok.FileName, fileId, "edit", token);
        try
        {
            StatInfo? statInfo = null;
            string? saveTime = null;
            if (Config.IsTrue("general", "detectexternalmodifications", true))
            {
                // check now the destination file against conflicts if required
                statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
                saveTime = statInfo.Xattrs.GetValueOrDefault(WopiUtils.LastSaveTimeKey);
                if (!IsValidSaveTime(saveTime, statInfo.MTime))
                {
                    // no xattr was there or we got our xattr but mtime is more recent: someone may have updated the file from
                    // a different source (e.g. FUSE or SMB mount), therefore force conflict and return failure to the application
                    log.LogWarning(
                        "Detected external modification, forcing conflict {user} {filename} {savetime} {mtime} {token}",
                        acctok.UserId.Short(),
                        acctok.FileName,
                        saveTime,
                        statInfo.MTime,
                        token);
                    return await utils.StoreAfterConflict(context, acctok, WopiUtils.ExternalLock, requestedLock, "The file being edited got moved or overwritten");
                }
            }

            // check whether to skip creating a new version of the file by reusing the same savetime attribute
            var noVersion = false;
            var noVersionGracePeriod = Config.GetInt("general", "noversiongraceperiod", 0);
            if (noVersionGracePeriod > 0)
            {
                if (statInfo == null)
                {
                    statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
                    saveTime = statInfo.Xattrs.GetValueOrDefault(WopiUtils.LastSaveTimeKey);
                }

                // disable versioning if the file was last saved more recently than noversiongraceperiod seconds ago
                noVersion = saveTime != null &&
                            saveTime.All(char.IsAsciiDigit) &&
                            long.TryParse(saveTime, CultureInfo.InvariantCulture, out var saved) &&
                            saved > state.Now - noVersionGracePeriod;
            }

            // Go for overwriting the file. Note that the entire check+write operation should be atomic,
            // but the previous checks still give the opportunity of a race condition. We just live with it.
            // Also, note we can't get a time resolution better than one second!
            // Anyhow, the EFSS should support versioning for such cases.
            await utils.StoreWopiFile(context, acctok, retrievedLock, WopiUtils.LastSaveTimeKey, noVersion: noVersion);
            statInfo = await Storage.StatX(acctok.Endpoint, acctok.FileName, acctok.UserId);
            log.LogInformation(
                "File stored successfully {action} {user} {filename} {version} {appname} {token}",
                "edit",
                acctok.UserId.Short(),
                acctok.FileName,
                statInfo.Etag,
                acctok.AppName,
                token);
            var response = new WopiResult(StatusCodes.Status200OK);
            response.Headers["X-WOPI-ItemVersion"] = $"v{statInfo.Etag}";
            return response;
        }
        catch (StorageException exception)
        {
            await utils.StoreForRecovery(acctok.WopiUser, acctok.FileName, token, exception.Message, null, context);
            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Implements the PutUserInfo WOPI call.
    /// </summary>
    public async Task<WopiResult> PutUserInfo(HttpContext context, string fileId, AccessToken acctok)
    {
        var token = context.TokenForLog();
        try
        {
            using var reader = new StreamReader(context.Request.Body, Encoding.UTF8);
            var userInfo = await reader.ReadToEndAsync();
            var storageLock = await Storage.GetLock(acctok.Endpoint, acctok.FileName, acctok.UserId);
            LockMetadata? lockMetadata = null;
            if (storageLock != null)
            {
                lockMetadata = new(acctok.AppName, storageLock.LockId);
            }

            await Storage.SetXattr(acctok.Endpoint, acctok.FileName, acctok.UserId, $"{WopiUtils.UserInfoKey}.{acctok.WopiUserName}", userInfo, lockMetadata);
            log.LogInformation("PutUserInfo {user} {filename} {fileid} {token}", acctok.UserId.Short(), acctok.FileName, fileId, token);
            return WopiResult.Message("OK", StatusCodes.Status200OK);
        }
        catch (StorageException exception)
        {
            log.LogError("PutUserInfo failed {filename} {error} {token}", acctok.FileName, exception.Message, token);
            return WopiResult.Message(ioError, StatusCodes.Status500InternalServerError);
        }
    }
}
