/// <summary>
/// The WOPI bridge extension. This connector service supports CodiMD and Etherpad.
/// </summary>
class Bridge(WopiState state, WopiUtils utils, BridgeHttp http, ILoggerFactory loggers)
{
    /// <summary>
    /// The supported plugins integrated with the WOPI Bridge extensions, by file extension.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string> ExtPlugins = new Dictionary<string, string>
    {
        ["md"] = "codimd",
        ["txt"] = "codimd",
        ["zmd"] = "codimd",
        ["epd"] = "etherpad",
        ["zep"] = "etherpad"
    };

    /// <summary>
    /// A header that bridged apps MUST send to the save endpoint to identify themselves.
    /// </summary>
    public const string BridgedAppNameHeader = "X-Efss-Bridged-App";

    /// <summary>
    /// A standard message to be displayed by the app when some content might be lost: this would only
    /// appear in case of uncaught exceptions or bugs handling the webhook callbacks.
    /// </summary>
    const string recoverMessage = "Please copy the content to a safe place and reopen the document again to paste it back.";

    readonly ILogger log = loggers.CreateLogger("bridge");
    readonly WopiClient wopi = new(http.Follow, state.Config.IsTrueOrYes("bridge", "sslverify", true), loggers.CreateLogger("wopiclient"));
    readonly ConcurrentDictionary<string, BridgeApp> plugins = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Guards <see cref="openFiles"/> and <see cref="saveResponses"/>: the save loop holds it for a whole round.
    /// </summary>
    readonly SemaphoreSlim gate = new(1, 1);
    readonly SemaphoreSlim wake = new(0);

    // a map of all open documents: wopisrc -> metadata, see OpenFile
    readonly Dictionary<string, OpenFile> openFiles = [];

    // a map of the latest responses to save requests, by wopisrc
    readonly Dictionary<string, BridgeReply> saveResponses = [];

    WopiConfig Config => state.Config;

    int SaveInterval => Config.GetInt("bridge", "saveinterval", 200);

    int UnlockInterval => Config.GetInt("bridge", "unlockinterval", 90);

    bool DisableZip => Config.IsTrueOrYes("bridge", "disablezip", false);

    long Now => state.Now;

    /// <summary>
    /// Loads the plugins of the apps configured in wopiserver.conf.
    /// </summary>
    public async Task Init()
    {
        foreach (var app in ExtPlugins.Values.Distinct())
        {
            var url = Config.Get("general", $"{app}url");
            if (string.IsNullOrEmpty(url))
            {
                continue;
            }

            var intUrl = Config.Get("general", $"{app}inturl");
            string? apiKey = null;
            var apiKeyFile = $"/var/run/secrets/{app}_apikey";
            if (File.Exists(apiKeyFile))
            {
                apiKey = (await File.ReadAllLinesAsync(apiKeyFile, Encoding.UTF8)).FirstOrDefault()?.TrimEnd('\n');
            }

            await LoadPlugin(app, url, string.IsNullOrEmpty(intUrl) ? url : intUrl, apiKey, "");
        }
    }

    /// <summary>
    /// Loads the plugin for the given app, if supported by the bridge service.
    /// </summary>
    async Task LoadPlugin(string appName, string appUrl, string appIntUrl, string? apiKey, string client)
    {
        var name = appName.ToLowerInvariant();
        if (plugins.TryGetValue(name, out var existing))
        {
            // already initialized, check that the app URL matches: the current model does not support multiple app backends
            if (appUrl != existing.AppExtUrl)
            {
                log.LogWarning("Attempt to use plugin with another appurl {client} {app} {appurl}", client, appName, appUrl);
                throw new PluginMismatchException(appName);
            }

            return;
        }

        if (!IsSupported(appName))
        {
            throw new UnsupportedAppException(appName);
        }

        var appLog = loggers.CreateLogger(name);
        BridgeApp plugin = name == "codimd" ? new CodiMD(http, wopi, appLog) : new Etherpad(http, wopi, appLog);
        plugin.AppName = appName;
        plugin.DisableZip = DisableZip;
        try
        {
            await plugin.Init(appUrl, appIntUrl, apiKey);
        }
        catch (Exception exception)
        {
            log.LogWarning("Failed to initialize plugin {app} {URL} {exception}", name, appIntUrl, exception.Message);
            throw new UnsupportedAppException(appName);
        }

        if (plugins.TryAdd(name, plugin))
        {
            log.LogInformation("Imported plugin for application {app} {plugin}", name, plugin.GetType().Name);
        }
    }

    /// <summary>
    /// Whether a given application is supported by one of the bridge extensions.
    /// </summary>
    public static bool IsSupported(string appName) =>
        ExtPlugins.Values.Contains(appName.ToLowerInvariant());

    /// <summary>
    /// Whether a given file extension is supported by one of the bridge extensions.
    /// </summary>
    public static bool IsExtSupported(string extension) =>
        ExtPlugins.ContainsKey(extension.ToLowerInvariant());

    /// <summary>
    /// The plugin whose appname matches (case-insensitively) the given appname.
    /// </summary>
    BridgeApp ValidateAppName(string appName)
    {
        foreach (var plugin in plugins.Values)
        {
            if (plugin.AppName.Contains(appName, StringComparison.OrdinalIgnoreCase))
            {
                return plugin;
            }
        }

        log.LogDebug("BridgeSave: unknown application {appname} {plugins}", appName, string.Join(", ", plugins.Keys));
        throw new UnsupportedAppException(appName);
    }

    /// <summary>
    /// Generates a URL safe hash of the wopisrc to be used as document id by the app.
    /// </summary>
    string GenDocId(string wopiSrc)
    {
        var digest = HMACSHA1.HashData(Encoding.UTF8.GetBytes(state.WopiSecret), Encoding.UTF8.GetBytes(wopiSrc.Split('/')[^1]));
        var docId = Convert.ToBase64String(digest).Replace('+', '-').Replace('/', '_')[..^1];
        log.LogDebug("Generating docid {wopisrc} {docid}", wopiSrc, docId);
        return docId;
    }

    // The Bridge endpoints start here

    /// <summary>
    /// Opens a doc by contacting the provided WOPISrc with the given access_token.
    /// Returns the app URL if successful, throws <see cref="FailedOpenException"/> otherwise.
    /// </summary>
    public async Task<string> AppOpen(string wopiSrc, string acctok, string appName, string appUrl, string appIntUrl, string? apiKey, ViewMode viewMode, string? revaToken, string client)
    {
        wopiSrc = UrlCodec.UnquotePlus(wopiSrc);

        // (re)load plugin and validate URLs
        BridgeApp app;
        try
        {
            await LoadPlugin(appName, appUrl, appIntUrl, apiKey, client);
            app = ValidateAppName(appName);
            log.LogDebug("BridgeOpen: processing supported app {appname} {plugin}", app.AppName, app.GetType().Name);
        }
        catch (UnsupportedAppException)
        {
            log.LogWarning("BridgeOpen: appname not supported or missing plugin {appname} {token}", appName, acctok.Short());
            throw new FailedOpenException($"Failed to load WOPI bridge plugin for {appName}", StatusCodes.Status500InternalServerError);
        }
        catch (PluginMismatchException)
        {
            log.LogError("BridgeOpen: app already configured {appname} {appurl} {token}", appName, appUrl, acctok.Short());
            throw new FailedOpenException($"Bridged app {appName} already configured with a different appurl", StatusCodes.Status501NotImplemented);
        }

        // WOPI GetFileInfo
        var response = await wopi.Request(wopiSrc, acctok, HttpMethod.Get);
        if (response.Status != StatusCodes.Status200OK)
        {
            log.LogWarning("BridgeOpen: unable to fetch file WOPI metadata {response}", response.Status);
            throw new FailedOpenException("Invalid WOPI context", StatusCodes.Status404NotFound);
        }

        var fileMetadata = response.Json()!.AsObject();
        try
        {
            BridgeLock wopiLock;
            // use the 'UserCanWrite' attribute to decide whether the file is to be opened in read-only mode
            if ((bool?) fileMetadata["UserCanWrite"] == true)
            {
                try
                {
                    // was it already being worked on?
                    wopiLock = await wopi.GetLock(wopiSrc, acctok);
                    log.LogInformation("Lock already held {lock} {token}", wopiLock.ToJson(), acctok.Short());
                    // add this token to the list, if not already in
                    if (!wopiLock.ToClose.ContainsKey(acctok.Short()))
                    {
                        wopiLock = await wopi.RefreshLock(wopiSrc, acctok, wopiLock);
                    }
                }
                catch (InvalidLockException exception)
                {
                    if (exception.Message != "404")
                    {
                        // lock is invalid/corrupted: force read-only mode
                        log.LogInformation("Invalid lock, forcing read-only mode {error} {token}", exception.Message, acctok.Short());
                        fileMetadata["UserCanWrite"] = false;
                    }

                    // otherwise, this is the first user opening the file; in both cases, fetch it
                    wopiLock = await app.LoadFromStorage(fileMetadata, wopiSrc, acctok, GenDocId(wopiSrc));
                    // and WOPI Lock it
                    response = await wopi.Request(
                        wopiSrc,
                        acctok,
                        HttpMethod.Post,
                        headers: new()
                        {
                            ["X-WOPI-Lock"] = wopiLock.ToJson(),
                            ["X-Wopi-Override"] = "LOCK"
                        });
                    if (response.Status != StatusCodes.Status200OK)
                    {
                        // failed to lock the file: open in read-only mode
                        log.LogWarning("Failed to lock the file {response} {token}", response.Status, acctok.Short());
                        fileMetadata["UserCanWrite"] = false;
                    }
                }

                await Track(wopiSrc, acctok, wopiLock, app);
            }
            else
            {
                // user has no write privileges, just fetch the document and push it to the app on a random docid
                wopiLock = await app.LoadFromStorage(fileMetadata, wopiSrc, acctok, null);
            }

            return await app.GetRedirectUrl(
                viewMode,
                wopiSrc,
                acctok,
                wopiLock.Doc[1..],
                FilePath(fileMetadata, acctok) + (string?) fileMetadata["BaseFileName"],
                (string?) fileMetadata["UserFriendlyName"] ?? "",
                revaToken);
        }
        catch (AppFailure exception)
        {
            // this can be raised by LoadFromStorage or GetRedirectUrl
            var message = exception.Message.Length > 0 ? exception.Message : "Unable to load the app, please try again later or contact support";
            throw new FailedOpenException(message, StatusCodes.Status500InternalServerError);
        }
    }

    /// <summary>
    /// Keeps track of an open document for the save loop and for statistical purposes.
    /// </summary>
    async Task Track(string wopiSrc, string acctok, BridgeLock wopiLock, BridgeApp app)
    {
        await gate.WaitAsync();
        try
        {
            if (openFiles.TryGetValue(wopiSrc, out var openFile))
            {
                // use the new acctok and the new/current wopilock content
                openFile.AccessToken = acctok;
                openFile.ToClose = new(wopiLock.ToClose);
            }
            else
            {
                openFiles[wopiSrc] = new()
                {
                    AccessToken = acctok,
                    ToSave = false,
                    LastSave = Now - SaveInterval,
                    ToClose = new()
                    {
                        [acctok.Short()] = false
                    },
                    DocId = wopiLock.Doc,
                    App = app.AppName
                };
            }

            // also clear any potential stale response for this document
            saveResponses.Remove(wopiSrc);
        }
        finally
        {
            gate.Release();
        }
    }

    /// <summary>
    /// Extracts the path from the folder URL. TODO this works with Reva master, not with Reva edge!
    /// </summary>
    string FilePath(JsonObject fileMetadata, string acctok)
    {
        var folderUrl = (string?) fileMetadata["BreadcrumbFolderUrl"];
        if (folderUrl == null)
        {
            return "";
        }

        if (!Uri.TryCreate(folderUrl, UriKind.Absolute, out var uri))
        {
            log.LogWarning("Failed to parse folderUrl, ignoring {url} {error} {token}", folderUrl, "invalid URL", acctok.Short());
            return "";
        }

        var path = Uri.UnescapeDataString(uri.AbsolutePath);
        if (path.StartsWith("/s/", StringComparison.Ordinal))
        {
            // top of public link, no leading /
            return $"{path[3..]}/";
        }

        if (path.StartsWith("/files/public/show/", StringComparison.Ordinal))
        {
            // subfolder of public link, no leading /
            return $"{path[19..]}/";
        }

        if (path.StartsWith("/files/spaces/", StringComparison.Ordinal))
        {
            // direct path to resource with leading /
            return $"{path[13..]}/";
        }

        // other folderurl structures are not supported for the time being
        return "";
    }

    /// <summary>
    /// Saves a doc given its WOPI context, and returns a JSON-formatted message. The actual save is asynchronous.
    /// </summary>
    public async Task<IResult> AppSave(HttpContext context, string docId)
    {
        var request = context.Request;
        string wopiSrc;
        string acctok;
        bool isClose;
        BridgeApp app;
        try
        {
            // fetch required metadata from request, return BAD_REQUEST if missing
            wopiSrc = UrlCodec.Unquote(request.Arg("WOPISrc") ?? throw new KeyNotFoundException("'WOPISrc'"));
            acctok = request.Arg("access_token") ?? throw new KeyNotFoundException("'access_token'");
            isClose = request.Arg("close") == "true";

            // ensure a save request comes from known/registered applications: this is done via a specific header
            app = ValidateAppName(request.Header(BridgedAppNameHeader) ?? throw new KeyNotFoundException($"'{BridgedAppNameHeader}'"));
            log.LogInformation("BridgeSave: requested action {isclose} {docid} {app} {wopisrc} {token}", isClose, docId, app.AppName, wopiSrc, acctok.Short());
        }
        catch (KeyNotFoundException exception)
        {
            log.LogError("BridgeSave: missing metadata {address} {headers} {args} {error}", context.RemoteAddress(), Headers(request), request.QueryString.Value, exception.Message);
            return ToResult(WopiClient.Reply($"Missing metadata, could not save. {recoverMessage}", StatusCodes.Status400BadRequest));
        }
        catch (UnsupportedAppException)
        {
            log.LogError("BridgeSave: unknown application {address} {appheader} {args}", context.RemoteAddress(), request.Header(BridgedAppNameHeader), request.QueryString.Value);
            return ToResult(WopiClient.Reply($"Unknown application, could not save. {recoverMessage}", StatusCodes.Status400BadRequest));
        }

        // enqueue the request, it will be processed asynchronously
        await gate.WaitAsync();
        try
        {
            // decide whether to notify the save loop
            var doNotify = isClose ||
                           !openFiles.TryGetValue(wopiSrc, out var openFile) ||
                           openFile.LastSave < Now - SaveInterval;
            if (openFiles.TryGetValue(wopiSrc, out openFile))
            {
                openFile.ToSave = true;
                openFile.ToClose[acctok.Short()] = isClose;
            }
            else
            {
                log.LogInformation("Save: repopulating missing metadata {wopisrc} {token}", wopiSrc, acctok.Short());
                openFiles[wopiSrc] = new()
                {
                    AccessToken = acctok,
                    ToSave = true,
                    LastSave = Now - SaveInterval,
                    ToClose = new()
                    {
                        [acctok.Short()] = isClose
                    },
                    DocId = docId,
                    App = app.AppName
                };
                // if it's the first time we heard about this wopisrc, remove any potential stale response
                saveResponses.Remove(wopiSrc);
            }

            if (doNotify &&
                wake.CurrentCount == 0)
            {
                // the save loop runs once this request releases the gate
                wake.Release();
            }

            // return latest known state for this document
            if (saveResponses.Remove(wopiSrc, out var reply))
            {
                if (reply.Status == StatusCodes.Status500InternalServerError)
                {
                    log.LogError("BridgeSave: returned response {response} {token}", reply.ToString(), acctok.Short());
                }
                else
                {
                    log.LogInformation("BridgeSave: returned response {response} {token}", reply.ToString(), acctok.Short());
                }

                return ToResult(reply);
            }

            log.LogInformation("BridgeSave: enqueued action {immediate} {token}", doNotify, acctok.Short());
            return WopiResult.Json("{}", StatusCodes.Status202Accepted);
        }
        finally
        {
            gate.Release();
        }
    }

    static string Headers(HttpRequest request) =>
        string.Join(", ", request.Headers.Select(_ => $"{_.Key}: {_.Value}"));

    static WopiResult ToResult(BridgeReply reply) =>
        WopiResult.Json(reply.Message, reply.Status);

    /// <summary>
    /// Returns the list of all currently opened files, for operators only.
    /// </summary>
    public async Task<IResult> AppList(HttpContext context)
    {
        // the apikey argument is added for convenience
        if (context.Request.Header("Authorization") != $"Bearer {state.WopiSecret}" &&
            context.Request.Arg("apikey") != state.WopiSecret)
        {
            log.LogWarning("BridgeList: unauthorized access attempt, missing authorization token {client}", context.RemoteAddress());
            return WopiResult.Text("Client not authorized", StatusCodes.Status401Unauthorized);
        }

        log.LogInformation("BridgeList: returning list of open files {client}", context.RemoteAddress());
        await gate.WaitAsync();
        try
        {
            return WopiResult.Json(Json.Serialize(openFiles));
        }
        finally
        {
            gate.Release();
        }
    }

    /// <summary>
    /// Validates that the given docid matches the wopiSrc.
    /// </summary>
    public bool ValidateDocId(string wopiSrc, string docId)
    {
        var expected = GenDocId(wopiSrc);
        log.LogDebug("Validate docid {expected} {actual} {wopisrc}", expected, docId, wopiSrc);
        return expected == docId;
    }

    // The save loop

    /// <summary>
    /// Waits until a save request asks for an immediate round, or the timeout elapses.
    /// </summary>
    public Task WaitForWork(TimeSpan timeout, Cancel cancel) =>
        wake.WaitAsync(timeout, cancel);

    /// <summary>
    /// Performs all pending save to storage operations.
    /// </summary>
    public async Task RunRound()
    {
        await gate.WaitAsync();
        try
        {
            // entries are eventually deleted from the dict
            foreach (var (wopiSrc, openFile) in openFiles.ToList())
            {
                try
                {
                    var wopiLock = await SaveDirty(openFile, wopiSrc);
                    wopiLock = await CloseWhenIdle(openFile, wopiSrc, wopiLock);
                    await Cleanup(openFile, wopiSrc, wopiLock);
                }
                catch (Exception exception)
                {
                    log.LogCritical("SaveThread: unexpected exception caught {ex} {type} {traceback}", exception.Message, exception.GetType().Name, exception.ToString());
                }
            }
        }
        finally
        {
            gate.Release();
        }
    }

    static bool Intersection(Dictionary<string, bool> values) =>
        values.Values.All(_ => _);

    static bool Union(Dictionary<string, bool> values) =>
        values.Values.Any(_ => _);

    /// <summary>
    /// Saves documents that are dirty for more than `saveinterval` or that are being closed.
    /// </summary>
    async Task<BridgeLock?> SaveDirty(OpenFile openFile, string wopiSrc)
    {
        if (!openFile.ToSave ||
            !Intersection(openFile.ToClose) && openFile.LastSave >= Now - SaveInterval)
        {
            return null;
        }

        var app = plugins[openFile.App];
        BridgeLock wopiLock;
        try
        {
            wopiLock = await wopi.GetLock(wopiSrc, openFile.AccessToken);
        }
        catch (InvalidLockException lockException)
        {
            if (lockException.Message == "401")
            {
                // this token has expired, nothing we can do any longer: by experience this happens on left-over
                // browser sessions, and the file was fully saved. Therefore just clean up by using some 'fake' metadata
                log.LogWarning("SaveThread: discarding file as token has expired {token} {docid}", openFile.AccessToken.Short(), openFile.DocId);
                Invalidate(openFile);
                return null;
            }

            log.LogInformation("SaveThread: attempting to relock file {token} {docid}", openFile.AccessToken.Short(), openFile.DocId);
            try
            {
                wopiLock = await wopi.Relock(wopiSrc, openFile.AccessToken, openFile.DocId, Intersection(openFile.ToClose));
            }
            catch (InvalidLockException relockException)
            {
                // even this attempt failed, give up
                saveResponses[wopiSrc] = WopiClient.Reply(relockException.Message, StatusCodes.Status500InternalServerError);
                // attempt to save to local storage to help for later recovery: this is a feature of the core wopiserver
                try
                {
                    var docLock = new BridgeLock
                    {
                        Doc = openFile.DocId,
                        FileName = "",
                        Digest = "dirty",
                        ToClose = []
                    };
                    var content = await app.Fetch(docLock, openFile.AccessToken);
                    await utils.StoreForRecovery("unknown", wopiSrc[(wopiSrc.LastIndexOf('/') + 1)..], openFile.AccessToken.Short(), relockException.Message, content, null);
                }
                catch (AppFailure)
                {
                    log.LogError(
                        "SaveThread: failed to fetch file for recovery to local storage {token} {docid} {app} {response}",
                        openFile.AccessToken.Short(),
                        openFile.DocId,
                        app.AppName,
                        StatusCodes.Status424FailedDependency);
                }

                // as above set some 'fake' metadata, will be automatically cleaned up later
                Invalidate(openFile);
                return null;
            }
        }

        // now save and log
        var reply = await app.SaveToStorage(wopiSrc, openFile.AccessToken, Intersection(openFile.ToClose), wopiLock);
        saveResponses[wopiSrc] = reply;
        openFile.LastSave = Now;
        if (reply.Status == StatusCodes.Status424FailedDependency)
        {
            // this is hopefully transient, yet we need to try until we get the file back to storage:
            // the updated lastsave time ensures next retry will happen after the saveinterval time
            if (!openFile.ToClose.ContainsKey("still-dirty"))
            {
                // add a special key that will prevent close/unlock and refresh lock. If the refresh fails,
                // the whole process will be retried at next round
                openFile.ToClose["still-dirty"] = false;
                wopiLock = await wopi.RefreshLock(wopiSrc, openFile.AccessToken, wopiLock, toClose: openFile.ToClose);
            }

            log.LogWarning(
                "SaveThread: failed to save, will retry {token} {docid} {lasterror} {tocl}",
                openFile.AccessToken.Short(),
                openFile.DocId,
                reply.ToString(),
                Json.Serialize(wopiLock.ToClose));
        }
        else
        {
            openFile.ToSave = false;
            // remove the special key above if present
            if (openFile.ToClose.Remove("still-dirty"))
            {
                wopiLock = await wopi.RefreshLock(wopiSrc, openFile.AccessToken, wopiLock, toClose: openFile.ToClose);
            }

            log.LogInformation(
                "SaveThread: file saved successfully {token} {docid} {tocl}",
                openFile.AccessToken.Short(),
                openFile.DocId,
                Json.Serialize(wopiLock.ToClose));
        }

        return wopiLock;
    }

    void Invalidate(OpenFile openFile)
    {
        openFile.LastSave = Now;
        openFile.ToSave = false;
        openFile.ToClose = new()
        {
            ["invalid-lock"] = true
        };
    }

    /// <summary>
    /// Closes and unlocks documents that are idle for more than 4x the save interval (about 14 minutes by default).
    /// They will transparently be relocked when/if the session resumes, but we seem to miss some close notifications,
    /// therefore this also works as a cleanup step.
    /// </summary>
    async Task<BridgeLock?> CloseWhenIdle(OpenFile openFile, string wopiSrc, BridgeLock? wopiLock)
    {
        if (openFile.LastSave >= Now - 4 * SaveInterval)
        {
            return wopiLock;
        }

        try
        {
            wopiLock ??= await wopi.GetLock(wopiSrc, openFile.AccessToken);
            // this will force a close in the cleanup step
            openFile.ToClose = openFile.ToClose.ToDictionary(_ => _.Key, _ => true);
            log.LogInformation("SaveThread: force-closing document {lastsavetime} {toclosetokens}", openFile.LastSave, Json.Serialize(openFile.ToClose));
        }
        catch (InvalidLockException)
        {
            // lock is gone, just cleanup our metadata
            log.LogWarning("SaveThread: cleaning up metadata, detected missed close event {url}", wopiSrc);
            openFiles.Remove(wopiSrc);
        }

        return wopiLock;
    }

    /// <summary>
    /// Removes state for closed documents after some time.
    /// </summary>
    async Task Cleanup(OpenFile openFile, string wopiSrc, BridgeLock? wopiLock)
    {
        if (!Union(openFile.ToClose) ||
            openFile.ToSave)
        {
            return;
        }

        // check lock
        try
        {
            wopiLock ??= await wopi.GetLock(wopiSrc, openFile.AccessToken);
        }
        catch (InvalidLockException)
        {
            // nothing to do here, this document may have been closed by another wopibridge
            if (openFile.LastSave < Now - UnlockInterval)
            {
                // yet clean up only after the unlockinterval time, cf. the InvalidLock handling in SaveDirty()
                log.LogInformation("SaveThread: cleaning up metadata, file already unlocked {url}", wopiSrc);
                openFiles.Remove(wopiSrc);
            }

            return;
        }

        // reconcile list of toclose tokens
        openFile.ToClose = wopiLock.ToClose.ToDictionary(_ => _.Key, _ => _.Value || openFile.ToClose.GetValueOrDefault(_.Key));
        if (Intersection(openFile.ToClose))
        {
            if (openFile.LastSave < Now - UnlockInterval)
            {
                // nobody is still on this document and some time has passed, unlock
                var response = await wopi.Request(
                    wopiSrc,
                    openFile.AccessToken,
                    HttpMethod.Post,
                    headers: new()
                    {
                        ["X-WOPI-Lock"] = wopiLock.ToJson(),
                        ["X-Wopi-Override"] = "UNLOCK"
                    });
                if (response.Status != StatusCodes.Status200OK)
                {
                    log.LogWarning("SaveThread: failed to unlock {lastsavetime} {token} {response}", openFile.LastSave, openFile.AccessToken.Short(), response.Status);
                }
                else
                {
                    log.LogInformation("SaveThread: unlocked document {lastsavetime} {token}", openFile.LastSave, openFile.AccessToken.Short());
                }

                openFiles.Remove(wopiSrc);
            }

            return;
        }

        if (openFile.ToClose.Count != wopiLock.ToClose.Count ||
            openFile.ToClose.Any(_ => !wopiLock.ToClose.TryGetValue(_.Key, out var value) || value != _.Value))
        {
            // some user still on it, refresh lock if the toclose part has changed
            try
            {
                await wopi.RefreshLock(wopiSrc, openFile.AccessToken, wopiLock, toClose: openFile.ToClose);
            }
            catch (InvalidLockException)
            {
                log.LogWarning("SaveThread: failed to refresh lock, will retry {url}", wopiSrc);
            }
        }
    }
}

/// <summary>
/// An open document, where <see cref="AccessToken"/> is one of the access tokens for the document, and
/// <see cref="ToClose"/> maps the short tokens (20 last chars) of all known tokens to whether they were closed.
/// </summary>
class OpenFile
{
    [JsonPropertyName("acctok")]
    public required string AccessToken { get; set; }

    [JsonPropertyName("tosave")]
    public required bool ToSave { get; set; }

    [JsonPropertyName("lastsave")]
    public required long LastSave { get; set; }

    [JsonPropertyName("toclose")]
    public required Dictionary<string, bool> ToClose { get; set; }

    [JsonPropertyName("docid")]
    public required string DocId { get; set; }

    [JsonPropertyName("app")]
    public required string App { get; set; }
}

/// <summary>
/// The open of a document through the bridge failed: the message and status are returned to the caller.
/// </summary>
class FailedOpenException(string message, int status) :
    Exception(message)
{
    public int Status { get; } = status;
}

class UnsupportedAppException(string appName) :
    Exception(appName);

class PluginMismatchException(string appName) :
    Exception(appName);

/// <summary>
/// The asynchronous save loop of the bridge.
/// </summary>
class BridgeSaveService(Bridge bridge, ILoggerFactory loggers) :
    BackgroundService
{
    readonly ILogger log = loggers.CreateLogger("bridge");

    protected override async Task ExecuteAsync(Cancel stoppingToken)
    {
        log.LogInformation("SaveThread starting");
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                // sleep for one minute or until awaken
                await bridge.WaitForWork(TimeSpan.FromSeconds(60), stoppingToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }

            await bridge.RunRound();
        }
    }

    public override Task StopAsync(Cancel cancellationToken)
    {
        log.LogInformation("Waiting for SaveThread to complete");
        return base.StopAsync(cancellationToken);
    }
}
