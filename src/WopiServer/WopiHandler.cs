using Microsoft.AspNetCore.Http.Features;

/// <summary>
/// The WOPI operations ([MS-WOPI] 3.3.5) over the <see cref="IWopiStorage"/>, the <see cref="IWopiLockStore"/> and the
/// <see cref="IWopiTokens"/>.
/// </summary>
sealed class WopiHandler(
    IWopiStorage storage,
    IWopiLockStore locks,
    IWopiTokens tokens,
    ProofKeys proofKeys,
    IOptions<WopiOptions> options,
    TimeProvider time,
    ILogger<WopiHandler> log,
    IServiceProvider services)
{
    /// <summary>
    /// The longest lock id, as advertised by SupportsExtendedLockLength.
    /// </summary>
    const int maxLockLength = 1024;

    /// <summary>
    /// The longest user info, as the WOPI documentation of PutUserInfo caps it.
    /// </summary>
    const int maxUserInfoLength = 1024;

    readonly IWopiUserInfoStore? userInfoStore = services.GetService<IWopiUserInfoStore>();

    WopiOptions Options => options.Value;

    /// <summary>
    /// The WOPISrc of a file.
    /// </summary>
    public static string FileUrl(string baseUrl, string fileId) =>
        $"{baseUrl}/wopi/files/{Uri.EscapeDataString(fileId)}";

    /// <summary>
    /// The access token of the request: the access_token parameter of [MS-WOPI] 3.3.5, or else the bearer token of the
    /// Authorization header of 2.2.1.
    /// </summary>
    static string? AccessToken(HttpRequest request)
    {
        var token = request.Query("access_token");
        if (!string.IsNullOrEmpty(token))
        {
            return token;
        }

        var authorization = request.Header("Authorization");
        if (authorization != null &&
            authorization.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            return authorization["Bearer ".Length..].Trim();
        }

        return null;
    }

    /// <summary>
    /// Validates the access token for the file, the X-WOPI-TimeStamp and the X-WOPI-Proof of the request.
    /// </summary>
    async Task<(WopiToken? Token, WopiResult? Error)> Authorize(HttpContext context, string fileId, string operation)
    {
        var request = context.Request;
        var value = AccessToken(request);
        WopiToken? token = null;
        if (value != null)
        {
            token = await tokens.Validate(value, context.RequestAborted);
        }

        if (token == null ||
            token.Access.FileId != fileId)
        {
            log.LogInformation("{Operation} of {FileId}: invalid or expired access token", operation, fileId);
            return (null, WopiResult.Unauthorized("Invalid access token"));
        }

        // a timestamp is only sent with a proof, but whenever present it must not be older than 20 minutes
        var timestamp = request.Header("X-WOPI-TimeStamp");
        if (timestamp != null &&
            (!long.TryParse(timestamp, NumberStyles.Integer, CultureInfo.InvariantCulture, out var ticks) ||
             ticks < time.GetUtcNow().UtcTicks - TimeSpan.FromMinutes(20).Ticks))
        {
            log.LogWarning("{Operation} of {FileId}: invalid or expired X-WOPI-TimeStamp", operation, fileId);
            // the WOPI validator expects a server error rather than an unauthorized request
            return (null, WopiResult.ServerError("Invalid or expired X-WOPI-TimeStamp"));
        }

        var url = $"{request.BaseUrl(Options)}{request.Path}{request.QueryString}";
        var proofError = await proofKeys.Validate(request, value!, url, context.RequestAborted);
        if (proofError != null)
        {
            log.LogWarning("{Operation} of {FileId}: invalid proof, {Reason}", operation, fileId, proofError);
            return (null, WopiResult.ServerError("Invalid or missing X-WOPI-Proof"));
        }

        log.LogDebug("{Operation} of {FileId} by {UserId}", operation, fileId, token.Access.UserId);
        return (token, null);
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.2 CheckFileInfo.
    /// </summary>
    public async Task<IResult> CheckFileInfo(HttpContext context, string fileId)
    {
        var (token, error) = await Authorize(context, fileId, "CheckFileInfo");
        if (error != null)
        {
            return error;
        }

        var cancel = context.RequestAborted;
        var file = await storage.GetFile(fileId, cancel);
        if (file == null)
        {
            return WopiResult.NotFound();
        }

        var access = token!.Access;
        var canWrite = access.CanWrite;
        var properties = new JsonObject
        {
            ["BaseFileName"] = file.Name,
            ["OwnerId"] = file.OwnerId,
            ["Size"] = file.Size,
            ["UserId"] = access.UserId,
            ["Version"] = file.Version,
            ["LastModifiedTime"] = file.LastModified.UtcDateTime.ToString("O", CultureInfo.InvariantCulture),
            ["IsAnonymousUser"] = access.IsAnonymous,
            ["SupportsExtendedLockLength"] = true,
            ["SupportsGetLock"] = true,
            ["SupportsLocks"] = canWrite,
            ["SupportsUpdate"] = canWrite,
            ["SupportsDeleteFile"] = canWrite,
            ["SupportsRename"] = true,
            ["UserCanWrite"] = canWrite,
            ["ReadOnly"] = !canWrite,
            ["UserCanRename"] = access.CanRename,
            ["UserCanNotWriteRelative"] = !access.CanWriteRelative,
            ["RestrictedWebViewOnly"] = access.ViewOnly,
            ["SupportsUserInfo"] = userInfoStore != null,
            // the same restriction for Collabora Online and ONLYOFFICE, which ignore RestrictedWebViewOnly
            ["DisableExport"] = access.ViewOnly,
            ["DisableCopy"] = access.ViewOnly,
            ["DisablePrint"] = access.ViewOnly,
            ["HideExportOption"] = access.ViewOnly
        };

        // [MS-WOPI] 3.3.5.1.2.2: the value MUST begin with a ".", so a file without extension has none
        var extension = FileNames.Extension(file.Name);
        if (extension.Length > 0)
        {
            properties["FileExtension"] = extension;
        }

        if (access.UserName != null)
        {
            properties["UserFriendlyName"] = access.UserName;
        }

        if (userInfoStore != null)
        {
            var userInfo = await userInfoStore.Get(access.UserId, cancel);
            if (!string.IsNullOrEmpty(userInfo))
            {
                properties["UserInfo"] = userInfo;
            }
        }

        AddHostUrls(properties, file, access);
        if (Options.OnCheckFileInfo != null)
        {
            await Options.OnCheckFileInfo(new(context, file, access, properties));
        }

        return WopiResult.Json(properties);
    }

    void AddHostUrls(JsonObject properties, WopiFile file, WopiAccess access)
    {
        var editUrl = Options.HostEditUrl?.Invoke(file, access);
        if (editUrl != null)
        {
            properties["HostEditUrl"] = editUrl;
        }

        var viewUrl = Options.HostViewUrl?.Invoke(file, access);
        if (viewUrl != null)
        {
            properties["HostViewUrl"] = viewUrl;
        }
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.3.1 GetFile.
    /// </summary>
    public async Task<IResult> GetFile(HttpContext context, string fileId)
    {
        var (_, error) = await Authorize(context, fileId, "GetFile");
        if (error != null)
        {
            return error;
        }

        var cancel = context.RequestAborted;
        try
        {
            var file = await storage.GetFile(fileId, cancel);
            if (file == null)
            {
                return WopiResult.NotFound();
            }

            // the client tells the largest file it expects
            if (long.TryParse(context.Request.Header("X-WOPI-MaxExpectedSize"), CultureInfo.InvariantCulture, out var maxExpectedSize) &&
                file.Size > maxExpectedSize)
            {
                return WopiResult.Text($"The file is larger than X-WOPI-MaxExpectedSize ({maxExpectedSize} bytes)", StatusCodes.Status412PreconditionFailed);
            }

            var result = new WopiResult(StatusCodes.Status200OK)
            {
                ContentType = "application/octet-stream",
                BodyStream = await storage.OpenRead(fileId, cancel)
            };
            result.Headers["Content-Disposition"] = $"attachment; filename*=UTF-8''{Uri.EscapeDataString(file.Name)}";
            result.Headers["X-WOPI-ItemVersion"] = file.Version;
            return result;
        }
        catch (FileNotFoundException)
        {
            // deleted meanwhile
            return WopiResult.NotFound();
        }
        catch (IOException exception)
        {
            log.LogError(exception, "GetFile of {FileId} failed", fileId);
            return WopiResult.ServerError("Failed to read the file");
        }
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.3.2 PutFile.
    /// </summary>
    public async Task<IResult> PutFile(HttpContext context, string fileId)
    {
        var (token, error) = await Authorize(context, fileId, "PutFile");
        if (error != null)
        {
            return error;
        }

        if (!token!.Access.CanWrite)
        {
            return WopiResult.Unauthorized("The access token does not allow writing");
        }

        var cancel = context.RequestAborted;
        var file = await storage.GetFile(fileId, cancel);
        if (file == null)
        {
            return WopiResult.NotFound();
        }

        var requestedLock = context.Request.Header("X-WOPI-Lock");
        var current = await locks.Get(fileId, cancel);
        if (current == null)
        {
            // an unlocked file can only be written while empty, as when a new file is created
            if (file.Size != 0)
            {
                return WopiResult.Conflict(null, "The file is not locked");
            }
        }
        else if (current.Id != requestedLock)
        {
            return WopiResult.Conflict(current.Id, "The file is locked with another lock");
        }

        try
        {
            AllowFileSize(context);
            var saved = await storage.Write(fileId, context.Request.Body, cancel);
            log.LogInformation("Saved {FileId}, version {Version}", fileId, saved.Version);
            var result = WopiResult.Ok();
            result.Headers["X-WOPI-ItemVersion"] = saved.Version;
            return result;
        }
        catch (FileNotFoundException)
        {
            return WopiResult.NotFound();
        }
        catch (IOException exception) when (exception is not BadHttpRequestException)
        {
            log.LogError(exception, "PutFile of {FileId} failed", fileId);
            return WopiResult.ServerError("Failed to write the file");
        }
    }

    /// <summary>
    /// Replaces the request body limit of the server with <see cref="WopiOptions.MaxFileSize"/>.
    /// </summary>
    void AllowFileSize(HttpContext context)
    {
        var feature = context.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (feature is { IsReadOnly: false })
        {
            feature.MaxRequestBodySize = Options.MaxFileSize;
        }
    }

    /// <summary>
    /// The POST operations on a file, told apart by their X-WOPI-Override header.
    /// </summary>
    public async Task<IResult> FilesPost(HttpContext context, string fileId)
    {
        var operation = context.Request.Header("X-WOPI-Override");
        if (operation == null)
        {
            return WopiResult.Text("Missing X-WOPI-Override", StatusCodes.Status400BadRequest);
        }

        var (token, error) = await Authorize(context, fileId, operation);
        if (error != null)
        {
            return error;
        }

        var access = token!.Access;
        switch (operation)
        {
            case "GET_LOCK":
                return await GetLock(context, fileId);
            case "PUT_RELATIVE":
                return await PutRelativeFile(context, token);
            case "PUT_USER_INFO":
                return await PutUserInfo(context, access);
            case "RENAME_FILE":
                return await RenameFile(context, access);
        }

        if (operation is "LOCK" or "REFRESH_LOCK" or "UNLOCK" or "DELETE" &&
            !access.CanWrite)
        {
            return WopiResult.Unauthorized("The access token does not allow writing");
        }

        switch (operation)
        {
            case "LOCK":
                return await Lock(context, fileId);
            case "REFRESH_LOCK":
                return await RefreshLock(context, fileId);
            case "UNLOCK":
                return await Unlock(context, fileId);
            case "DELETE":
                return await DeleteFile(context, fileId);
        }

        return WopiResult.NotImplemented($"Unsupported operation {operation}");
    }

    /// <summary>
    /// The X-WOPI-Lock of a request, or a 400 when it is missing or too long.
    /// </summary>
    static (string? Lock, WopiResult? Error) RequestedLock(HttpContext context)
    {
        var requested = context.Request.Header("X-WOPI-Lock");
        string reason;
        if (string.IsNullOrEmpty(requested))
        {
            reason = "Missing X-WOPI-Lock";
        }
        else if (requested.Length > maxLockLength)
        {
            reason = $"X-WOPI-Lock is longer than {maxLockLength} characters";
        }
        else
        {
            return (requested, null);
        }

        var result = WopiResult.Text(reason, StatusCodes.Status400BadRequest);
        result.Headers["X-WOPI-LockFailureReason"] = reason;
        return (null, result);
    }

    WopiLock NewLock(string id) =>
        new(id, time.GetUtcNow() + Options.LockDuration);

    static WopiResult Locked(WopiFile file)
    {
        var result = WopiResult.Ok();
        result.Headers["X-WOPI-ItemVersion"] = file.Version;
        return result;
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.13 Lock, and 3.3.5.1.23 UnlockAndRelock when X-WOPI-OldLock is given.
    /// </summary>
    async Task<IResult> Lock(HttpContext context, string fileId)
    {
        var (requested, error) = RequestedLock(context);
        if (error != null)
        {
            return error;
        }

        var cancel = context.RequestAborted;
        var file = await storage.GetFile(fileId, cancel);
        if (file == null)
        {
            return WopiResult.NotFound();
        }

        var oldLock = context.Request.Header("X-WOPI-OldLock");
        if (!string.IsNullOrEmpty(oldLock))
        {
            var previous = await locks.CompareExchange(fileId, oldLock, NewLock(requested!), cancel);
            if (previous?.Id != oldLock)
            {
                return WopiResult.Conflict(previous?.Id, "The file is not locked with X-WOPI-OldLock");
            }

            log.LogInformation("Relocked {FileId}", fileId);
            return Locked(file);
        }

        var current = await locks.CompareExchange(fileId, null, NewLock(requested!), cancel);
        if (current == null)
        {
            log.LogInformation("Locked {FileId}", fileId);
            return Locked(file);
        }

        // locking again with the same lock refreshes it
        if (current.Id == requested)
        {
            current = await locks.CompareExchange(fileId, requested, NewLock(requested), cancel);
            if (current?.Id == requested)
            {
                return Locked(file);
            }
        }

        return WopiResult.Conflict(current?.Id, "The file is locked with another lock");
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.17 RefreshLock.
    /// </summary>
    async Task<IResult> RefreshLock(HttpContext context, string fileId)
    {
        var (requested, error) = RequestedLock(context);
        if (error != null)
        {
            return error;
        }

        var cancel = context.RequestAborted;
        if (await storage.GetFile(fileId, cancel) == null)
        {
            return WopiResult.NotFound();
        }

        var current = await locks.CompareExchange(fileId, requested, NewLock(requested!), cancel);
        if (current?.Id != requested)
        {
            return WopiResult.Conflict(current?.Id, "The file is not locked with X-WOPI-Lock");
        }

        return WopiResult.Ok();
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.22 Unlock.
    /// </summary>
    async Task<IResult> Unlock(HttpContext context, string fileId)
    {
        var (requested, error) = RequestedLock(context);
        if (error != null)
        {
            return error;
        }

        var cancel = context.RequestAborted;
        var file = await storage.GetFile(fileId, cancel);
        if (file == null)
        {
            return WopiResult.NotFound();
        }

        var current = await locks.CompareExchange(fileId, requested, null, cancel);
        if (current?.Id != requested)
        {
            return WopiResult.Conflict(current?.Id, "The file is not locked with X-WOPI-Lock");
        }

        log.LogInformation("Unlocked {FileId}", fileId);
        return Locked(file);
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.10 GetLock: an unlocked file has an empty lock.
    /// </summary>
    async Task<IResult> GetLock(HttpContext context, string fileId)
    {
        var cancel = context.RequestAborted;
        if (await storage.GetFile(fileId, cancel) == null)
        {
            return WopiResult.NotFound();
        }

        var current = await locks.Get(fileId, cancel);
        var result = WopiResult.Ok();
        result.Headers["X-WOPI-Lock"] = current?.Id ?? "";
        return result;
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.9 DeleteFile.
    /// </summary>
    async Task<IResult> DeleteFile(HttpContext context, string fileId)
    {
        var cancel = context.RequestAborted;
        var current = await locks.Get(fileId, cancel);
        if (current != null)
        {
            return WopiResult.Conflict(current.Id, "A locked file cannot be deleted");
        }

        try
        {
            await storage.Delete(fileId, cancel);
            log.LogInformation("Deleted {FileId}", fileId);
            return WopiResult.Ok();
        }
        catch (FileNotFoundException)
        {
            return WopiResult.NotFound();
        }
        catch (IOException exception)
        {
            log.LogError(exception, "DeleteFile of {FileId} failed", fileId);
            return WopiResult.ServerError("Failed to delete the file");
        }
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.21 RenameFile. The requested name comes without the extension, which the file keeps.
    /// </summary>
    async Task<IResult> RenameFile(HttpContext context, WopiAccess access)
    {
        var fileId = access.FileId;
        if (!access.CanRename)
        {
            return WopiResult.Unauthorized("The access token does not allow renaming");
        }

        var requestedName = context.Request.Header("X-WOPI-RequestedName");
        if (requestedName == null)
        {
            return WopiResult.Text("Missing X-WOPI-RequestedName", StatusCodes.Status400BadRequest);
        }

        var name = FileNames.DecodeUtf7(requestedName);
        if (!FileNames.IsLegal(name))
        {
            return InvalidFileName("The requested name is illegal");
        }

        var cancel = context.RequestAborted;
        var current = await locks.Get(fileId, cancel);
        var requestedLock = context.Request.Header("X-WOPI-Lock");
        if (current != null &&
            current.Id != requestedLock)
        {
            return WopiResult.Conflict(current.Id, "The file is locked with another lock");
        }

        try
        {
            var file = await storage.GetFile(fileId, cancel);
            if (file == null)
            {
                return WopiResult.NotFound();
            }

            var newName = $"{name}{FileNames.Extension(file.Name)}";
            if (Encoding.UTF8.GetByteCount(newName) > FileNames.MaxBytes)
            {
                return InvalidFileName("The name is too long once the extension is added");
            }

            if (await storage.Rename(fileId, newName, cancel) == null)
            {
                // a rename must not overwrite another file
                return InvalidFileName("A file with the requested name already exists");
            }

            log.LogInformation("Renamed {FileId} to {Name}", fileId, newName);
            return WopiResult.Json(
                new JsonObject
                {
                    ["Name"] = name
                });
        }
        catch (FileNotFoundException)
        {
            return WopiResult.NotFound();
        }
        catch (IOException exception)
        {
            log.LogError(exception, "RenameFile of {FileId} failed", fileId);
            return WopiResult.ServerError("Failed to rename the file");
        }
    }

    static WopiResult InvalidFileName(string reason)
    {
        var result = WopiResult.Text(reason, StatusCodes.Status400BadRequest);
        result.Headers["X-WOPI-InvalidFileNameError"] = reason;
        return result;
    }

    /// <summary>
    /// [MS-WOPI] 3.3.5.1.15 PutRelativeFile, the Save As of the WOPI clients: X-WOPI-SuggestedTarget is a name or an
    /// extension that can be changed to be legal and new, X-WOPI-RelativeTarget a name to use as is.
    /// </summary>
    async Task<IResult> PutRelativeFile(HttpContext context, WopiToken token)
    {
        var access = token.Access;
        var fileId = access.FileId;
        if (!access.CanWriteRelative)
        {
            // the WOPI validator expects NOT_IMPLEMENTED for a user who cannot write relative
            return WopiResult.NotImplemented("The access token does not allow creating files");
        }

        var request = context.Request;
        var suggested = request.Header("X-WOPI-SuggestedTarget");
        var relative = request.Header("X-WOPI-RelativeTarget");
        if (string.IsNullOrEmpty(suggested) == string.IsNullOrEmpty(relative))
        {
            return WopiResult.Text("Either X-WOPI-SuggestedTarget or X-WOPI-RelativeTarget is required", StatusCodes.Status400BadRequest);
        }

        var cancel = context.RequestAborted;
        try
        {
            var file = await storage.GetFile(fileId, cancel);
            if (file == null)
            {
                return WopiResult.NotFound();
            }

            AllowFileSize(context);
            // the content may be written more than once, when another file takes the name meanwhile
            request.EnableBuffering();
            WopiFile? created;
            if (!string.IsNullOrEmpty(suggested))
            {
                created = await CreateSuggested(request, file, FileNames.DecodeUtf7(suggested), cancel);
                if (created == null)
                {
                    return WopiResult.ServerError("No free name for the suggested target");
                }
            }
            else
            {
                var name = FileNames.DecodeUtf7(relative!);
                var overwrite = string.Equals(request.Header("X-WOPI-OverwriteRelativeTarget"), "true", StringComparison.OrdinalIgnoreCase);
                var (result, relativeFile) = await CreateRelative(token, request, name, overwrite, cancel);
                if (result != null)
                {
                    return result;
                }

                created = relativeFile!;
            }

            log.LogInformation("Created {FileId} {Name} from {Source}", created.Id, created.Name, fileId);
            var newToken = await tokens.Issue(access with { FileId = created.Id }, cancel);
            var properties = new JsonObject
            {
                ["Name"] = created.Name,
                ["Url"] = $"{FileUrl(request.BaseUrl(Options), created.Id)}?access_token={Uri.EscapeDataString(newToken.Value)}"
            };
            AddHostUrls(properties, created, newToken.Access);
            return WopiResult.Json(properties);
        }
        catch (FileNotFoundException)
        {
            return WopiResult.NotFound();
        }
        catch (UnauthorizedAccessException)
        {
            return WopiResult.NotImplemented("The storage does not allow creating files here");
        }
        catch (IOException exception) when (exception is not BadHttpRequestException)
        {
            log.LogError(exception, "PutRelativeFile of {FileId} failed", fileId);
            return WopiResult.ServerError("Failed to create the file");
        }
    }

    /// <summary>
    /// Creates the file under the suggested name, made legal and new, or null when no free name is found.
    /// </summary>
    async Task<WopiFile?> CreateSuggested(HttpRequest request, WopiFile file, string suggested, Cancel cancel)
    {
        string name;
        if (suggested.StartsWith('.'))
        {
            // only the extension
            name = FileNames.Legal($"{FileNames.SplitExtension(file.Name).Root}{suggested}");
        }
        else
        {
            name = FileNames.Legal(suggested);
        }

        for (var attempt = 0; attempt < 100; attempt++)
        {
            var free = await FreeName(file.Id, name, cancel);
            if (free == null)
            {
                return null;
            }

            request.Body.Position = 0;
            var created = await storage.CreateSibling(file.Id, free, request.Body, false, cancel);
            if (created != null)
            {
                return created;
            }
        }

        return null;
    }

    /// <summary>
    /// Creates the file under the given name, failing when it is illegal, or exists and is locked or not to be
    /// overwritten.
    /// </summary>
    async Task<(WopiResult? Error, WopiFile? File)> CreateRelative(WopiToken token, HttpRequest request, string name, bool overwrite, Cancel cancel)
    {
        var fileId = token.Access.FileId;
        if (!FileNames.IsLegal(name))
        {
            // the name must not be changed, so an illegal one fails the request, suggesting a legal one
            var illegal = WopiResult.Text("The relative target is illegal", StatusCodes.Status400BadRequest);
            await AddValidTarget(illegal, fileId, FileNames.Legal(name), cancel);
            return (illegal, null);
        }

        var existing = await storage.GetSibling(fileId, name, cancel);
        if (existing != null)
        {
            var existingLock = await locks.Get(existing.Id, cancel);
            if (!overwrite ||
                existingLock != null)
            {
                var conflict = WopiResult.Json(
                    new JsonObject
                    {
                        ["Name"] = existing.Name,
                        ["Url"] = FileUrl(request.BaseUrl(Options), existing.Id)
                    },
                    StatusCodes.Status409Conflict);
                conflict.Headers["X-WOPI-Lock"] = existingLock?.Id ?? "";
                conflict.Headers["X-WOPI-LockFailureReason"] = "The relative target exists";
                await AddValidTarget(conflict, fileId, name, cancel);
                return (conflict, null);
            }
        }

        var created = await storage.CreateSibling(fileId, name, request.Body, true, cancel);
        return (null, created);
    }

    /// <summary>
    /// Adds the X-WOPI-ValidRelativeTarget a client can use instead.
    /// </summary>
    async Task AddValidTarget(WopiResult result, string fileId, string name, Cancel cancel)
    {
        var free = await FreeName(fileId, name, cancel);
        if (free != null)
        {
            result.Headers["X-WOPI-ValidRelativeTarget"] = FileNames.EncodeUtf7(free);
        }
    }

    /// <summary>
    /// The name, or the first of `name (2).ext`, `name (3).ext` and so on that is free in the folder of the file.
    /// </summary>
    async Task<string?> FreeName(string fileId, string name, Cancel cancel)
    {
        var candidate = name;
        for (var number = 2; number < 1000; number++)
        {
            if (await storage.GetSibling(fileId, candidate, cancel) == null)
            {
                return candidate;
            }

            candidate = FileNames.Numbered(name, number);
        }

        return null;
    }

    /// <summary>
    /// The WOPI documentation of PutUserInfo: stores the user info, which CheckFileInfo returns as UserInfo.
    /// </summary>
    async Task<IResult> PutUserInfo(HttpContext context, WopiAccess access)
    {
        if (userInfoStore == null)
        {
            return WopiResult.NotImplemented("PutUserInfo is not supported");
        }

        var cancel = context.RequestAborted;
        using var reader = new StreamReader(context.Request.Body, Encoding.UTF8);
        var userInfo = await reader.ReadToEndAsync(cancel);
        if (userInfo.Length > maxUserInfoLength)
        {
            return WopiResult.Text($"The user info is longer than {maxUserInfoLength} characters", StatusCodes.Status400BadRequest);
        }

        if (await storage.GetFile(access.FileId, cancel) == null)
        {
            return WopiResult.NotFound();
        }

        await userInfoStore.Set(access.UserId, userInfo, cancel);
        return WopiResult.Ok();
    }
}
