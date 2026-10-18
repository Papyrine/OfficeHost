using System.Runtime.InteropServices;
using Prometheus;

/// <summary>
/// The HTTP endpoints of the server: the IOP endpoints, the WOPI protocol, and the bridge.
/// </summary>
class WopiEndpoints(WopiState state, WopiUtils utils, WopiCore core, Bridge bridge, ILogger log)
{
    static readonly Counter openByApp = Metrics.CreateCounter("open_by_app", "Number of /open calls by appname", "open_type");

    static WopiResult Unauthorized() =>
        WopiResult.Text("Client not authorized", StatusCodes.Status401Unauthorized);

    /// <summary>
    /// A handler taking only the context would bind as a RequestDelegate, which discards the result.
    /// </summary>
    static Delegate Handler(Func<HttpContext, Task<IResult>> handler) =>
        handler;

    bool IsIopAuthorized(HttpContext context) =>
        context.Request.Header("Authorization") == $"Bearer {state.IopSecret}";

    public void Map(WebApplication app)
    {
        app.Use(HandleException);
        app.Use(AddWopiHeaders);
        app.UseHttpMetrics();
        app.MapMetrics();

        app.MapGet("/", () => Results.Redirect("/wopi"));
        app.MapGet("/wopi", Handler(Index));

        // IOP endpoints
        app.MapGet("/wopi/iop/openinapp", Handler(OpenInApp));
        app.MapGet("/wopi/iop/download", Handler(Download));
        app.MapGet("/wopi/iop/list", GetOpenFiles);
        app.MapGet("/wopi/iop/conflicts", GetConflicts);
        app.MapGet("/wopi/iop/test", Handler(WopiTest));

        // WOPI protocol implementation
        app.MapGet("/wopi/files/{fileid}", CheckFileInfo);
        app.MapGet("/wopi/files/{fileid}/contents", GetFile);
        app.MapPost("/wopi/files/{fileid}", FilesPost);
        app.MapPost("/wopi/files/{fileid}/contents", PutFile);

        // Bridge functionality
        app.MapMethods("/wopi/bridge/{docid}", ["HEAD", "POST"], BridgeRoot);
        app.MapGet("/wopi/bridge/list", Handler(bridge.AppList));
    }

    /// <summary>
    /// Logs any uncaught exception and returns HTTP 500.
    /// </summary>
    async Task HandleException(HttpContext context, RequestDelegate next)
    {
        try
        {
            await next(context);
        }
        catch (Exception exception) when (!context.Response.HasStarted && !context.RequestAborted.IsCancellationRequested)
        {
            log.LogCritical(
                "Unexpected exception caught {type} {traceback} {client} {requestedUrl}",
                exception.GetType().FullName,
                exception.ToString(),
                context.Client(),
                context.Request.Url());
            context.Response.Clear();
            await WopiResult.Text("Internal error, please contact support", StatusCodes.Status500InternalServerError).ExecuteAsync(context);
        }
    }

    /// <summary>
    /// [MS-WOPI] 2.2.1 requires X-WOPI-ServerVersion and X-WOPI-MachineName on all WOPI responses, and
    /// X-WOPI-ServerError on 500 responses.
    /// </summary>
    static Task AddWopiHeaders(HttpContext context, RequestDelegate next)
    {
        if (context.Request.Path.StartsWithSegments("/wopi/files") ||
            context.Request.Path.StartsWithSegments("/wopi/folders"))
        {
            context.Response.OnStarting(
                () =>
                {
                    var headers = context.Response.Headers;
                    headers["X-WOPI-ServerVersion"] = WopiState.Version;
                    headers["X-WOPI-MachineName"] = Environment.MachineName;
                    if (context.Response.StatusCode == StatusCodes.Status500InternalServerError &&
                        !headers.ContainsKey("X-WOPI-ServerError"))
                    {
                        headers["X-WOPI-ServerError"] = "Server error";
                    }

                    return Task.CompletedTask;
                });
        }

        return next(context);
    }

    /// <summary>
    /// A default index page with some user-friendly information about this service.
    /// </summary>
    async Task<IResult> Index(HttpContext context)
    {
        log.LogDebug("Accessed index page {client}", context.RemoteAddress());
        var health = await state.Storage.HealthCheck();
        var html = $"""

                          <html><head><title>ScienceMesh WOPI Server</title></head>
                          <body>
                          <div align="center" style="color:#000080; padding-top:50px; font-family:Verdana; size:11">
                          This is the ScienceMesh IOP <a href=https://github.com/cs3org/wopiserver#readme>WOPI</a> server
                          to support online office-like editors.<br>
                          The service includes support for non-WOPI-native apps through a bridge extension.<br>
                          To use this service, please log in to your EFSS Storage and click on a supported document.</div>
                          <div style="position: absolute; bottom: 10px; left: 10px; width: 99%;"><hr>
                          <i>ScienceMesh WOPI Server {WopiState.Version} at {HostName()}. Powered by ASP.NET Core for {RuntimeInformation.FrameworkDescription}.
                             Storage type: <span style="font-family:monospace">{state.StorageType}</span>.
                             Health status: <span style="font-family:monospace">{health}</span>.</i>
                          </body>
                          </html>

                    """;
        var response = WopiResult.Text(html);
        response.Headers["X-Frame-Options"] = "sameorigin";
        response.Headers["X-XSS-Protection"] = "1; mode=block";
        return response;
    }

    static string HostName()
    {
        try
        {
            return Dns.GetHostEntry(Dns.GetHostName()).HostName;
        }
        catch (Exception exception) when (exception is System.Net.Sockets.SocketException or ArgumentException)
        {
            return Environment.MachineName;
        }
    }

    // IOP endpoints

    /// <summary>
    /// Generates a WOPISrc target and an access token to be passed to a WOPI-compatible Office-like app
    /// for accessing a given file for a given user.
    /// Headers:
    /// - Authorization: a bearer shared secret to protect this call as it provides direct access to any user's file.
    ///   This can be omitted if the storage is based on CS3, as Reva would authenticate calls via the TokenHeader below.
    /// - TokenHeader: an x-access-token to serve as user identity towards Reva
    /// - ApiKey (optional): a shared secret to be used with the end-user application if required
    /// - X-Trace-Id (optional): a trace id to cross-reference logs
    /// Request arguments:
    /// - enum viewmode: how the user should access the file, according to ViewMode/the CS3 app provider API
    /// - string fileid: the Reva fileid of the file to be opened
    /// - string endpoint (optional): the storage endpoint to be used to look up the file or the storage id, in case of
    ///   multi-instance underlying storage; defaults to 'default'
    /// - string username (optional): user's full display name, typically shown by the app; defaults to
    ///   'Guest ' + 3 random letters to represent anonymous users
    /// - string userid (optional): an unique identifier for the user, used internally by the app; defaults to
    ///   a random string of 10 characters to represent anonymous users
    /// - string folderurl (optional): the URL to come back to the containing folder for this file, typically shown by the app
    /// - string appname: the identifier of the end-user application to be served
    /// - string appurl: the URL of the end-user application
    /// - string appviewurl (optional): the URL of the end-user application in view mode when different (defaults to appurl)
    /// - string appinturl (optional): the internal URL of the end-user application (applicable with containerized deployments)
    /// - string usertype (optional): one of "regular", "federated", "ocm", "anonymous". Defaults to "regular"
    /// Returns a JSON response as follows:
    /// { "app-url" : "&lt;URL of the target application with query parameters&gt;", "form-parameters" : { "access_token" : "&lt;WOPI access token&gt;" } }
    /// or a message and a 4xx/5xx HTTP code in case of errors.
    /// </summary>
    async Task<IResult> OpenInApp(HttpContext context)
    {
        var request = context.Request;
        openByApp.WithLabels(request.Arg("appname") ?? "None").Inc();

        // validate tokens
        if (state.StorageType != "cs3" &&
            !IsIopAuthorized(context))
        {
            log.LogWarning("iopOpenInApp: unauthorized access attempt, missing authorization token {client} {clientAuth}", context.RemoteAddress(), request.Header("Authorization") ?? "None");
            return Unauthorized();
        }

        var userToken = request.Header("TokenHeader");
        if (userToken == null)
        {
            log.LogWarning("iopOpenInApp: missing TokenHeader in request {client}", context.RemoteAddress());
            return Unauthorized();
        }

        // validate all parameters
        var fileId = request.Arg("fileid") ?? "";
        if (fileId.Length == 0)
        {
            log.LogWarning("iopOpenInApp: fileid must be provided {client}", context.RemoteAddress());
            return WopiResult.Text("Missing fileid argument", StatusCodes.Status400BadRequest);
        }

        var viewModeArg = request.Arg("viewmode");
        if (WopiEnums.ParseViewMode(viewModeArg) is not { } viewMode)
        {
            log.LogWarning("iopOpenInApp: invalid viewmode parameter {client} {viewmode} {error}", context.RemoteAddress(), viewModeArg, $"'{viewModeArg}' is not a valid ViewMode");
            return WopiResult.Text("Missing or invalid viewmode argument", StatusCodes.Status400BadRequest);
        }

        var userName = UrlCodec.UnquotePlus(request.Arg("username") ?? "");
        // this needs to be a unique identifier: if missing (case of anonymous users), just generate a random string
        var wopiUser = request.Arg("userid") ?? WopiUtils.RandomString(10);
        // defaults to `/`
        var folderUrl = UrlCodec.UnquotePlus(request.Arg("folderurl") ?? "%2F");
        var endpoint = request.Arg("endpoint") ?? "default";
        var appName = UrlCodec.UnquotePlus(request.Arg("appname") ?? "");
        var appUrl = UrlCodec.UnquotePlus(request.Arg("appurl") ?? "").Trim('/');
        var appViewUrl = UrlCodec.UnquotePlus(request.Arg("appviewurl") ?? appUrl).Trim('/');
        var userTypeArg = request.Arg("usertype");
        var userType = UserType.Regular;
        if (userTypeArg != null)
        {
            if (WopiEnums.ParseUserType(userTypeArg) is { } parsed)
            {
                userType = parsed;
            }
            else
            {
                log.LogWarning("iopOpenInApp: invalid usertype, falling back to regular {client} {usertype} {error}", context.RemoteAddress(), userTypeArg, $"'{userTypeArg}' is not a valid UserType");
            }
        }

        if (appName.Length == 0 ||
            appUrl.Length == 0)
        {
            log.LogWarning("iopOpenInApp: app-related arguments must be provided {client}", context.RemoteAddress());
            return WopiResult.Text("Missing appname or appurl arguments", StatusCodes.Status400BadRequest);
        }

        var trace = request.Header("X-Trace-Id") ?? "N/A";
        GeneratedToken generated;
        try
        {
            (var userId, wopiUser) = state.Storage.GetUserIdFromCreds(userToken, wopiUser);
            generated = await utils.GenerateAccessToken(userId, fileId, viewMode, userName, wopiUser, userType, folderUrl, endpoint, appName, appUrl, appViewUrl, trace);
        }
        catch (StorageException exception)
        {
            log.LogInformation(
                "iopOpenInApp: remote error on generating token {client} {trace} {user} {friendlyname} {mode} {endpoint} {reason}",
                context.RemoteAddress(),
                trace,
                userToken.Short(),
                userName,
                viewMode.Value(),
                endpoint,
                exception.Message);
            return WopiResult.Text("Remote error, file not found or file is a directory", StatusCodes.Status404NotFound);
        }

        var result = new JsonObject();
        if (Bridge.IsSupported(appName))
        {
            try
            {
                result["app-url"] = await bridge.AppOpen(
                    utils.GenerateWopiSrc(generated.Inode),
                    generated.AccessToken,
                    appName,
                    appUrl,
                    UrlCodec.UnquotePlus(request.Arg("appinturl") ?? appUrl),
                    request.Header("ApiKey"),
                    generated.ViewMode,
                    userToken,
                    context.RemoteAddress());
                result["form-parameters"] = new JsonObject();
            }
            catch (FailedOpenException exception)
            {
                return WopiResult.Text(exception.Message, exception.Status);
            }
        }
        else
        {
            // the base app URL is the editor in READ_WRITE mode, and the viewer in READ_ONLY, PREVIEW or EMBEDDED mode
            // as the known WOPI applications all support switching from preview to edit mode
            var url = generated.ViewMode == ViewMode.ReadWrite ? appUrl : appViewUrl;
            url += $"{(url.Contains('?') ? '&' : '?')}WOPISrc={utils.GenerateWopiSrc(generated.Inode, appName == state.ProxiedAppName)}";
            if (state.Config.IsTrue("apps", "businessflow", false))
            {
                // tells the app to enable the business flow if appropriate
                url += "&IsLicensedUser=1";
            }

            result["app-url"] = url;
            result["form-parameters"] = new JsonObject
            {
                ["access_token"] = generated.AccessToken
            };
            if (generated.ViewMode != viewMode)
            {
                result["forced-viewmode-reason"] = generated.Details;
            }
        }

        var appForLog = (string) result["app-url"]!;
        var access = appForLog.IndexOf("access", StringComparison.Ordinal);
        if (access > 0)
        {
            appForLog = $"{appForLog[..access]}access_token=redacted";
        }

        log.LogInformation("iopOpenInApp: redirecting client {appurl}", appForLog);
        return WopiResult.Json(result);
    }

    /// <summary>
    /// Decodes and validates the access token of the request, without the issuer check of the WOPI calls.
    /// </summary>
    AccessToken? DecodeToken(HttpContext context, out string error)
    {
        error = "";
        try
        {
            var acctok = Jwt.Decode<AccessToken>(context.Request.Arg("access_token") ?? throw new JwtException("'access_token'"), state.WopiSecret, state.Now);
            if (acctok.Exp < state.Now)
            {
                throw new JwtException("Signature has expired", expired: true);
            }

            return acctok;
        }
        catch (JwtException exception)
        {
            error = exception.Message;
            return null;
        }
    }

    /// <summary>
    /// Returns the file's content for a given valid access token. Used as a download URL,
    /// so that the path and possibly the x-access-token are never explicitly visible.
    /// </summary>
    async Task<IResult> Download(HttpContext context)
    {
        var acctok = DecodeToken(context, out var error);
        if (acctok == null)
        {
            log.LogInformation("Expired or malformed token {client} {requestedUrl} {error} {token}", context.RemoteAddress(), context.Request.BaseUrl(), error, context.Request.Arg("access_token") ?? "N/A");
            return WopiResult.Text("Missing or invalid access token", StatusCodes.Status401Unauthorized);
        }

        log.LogInformation("iopDownload: returning contents {client} {endpoint} {filename} {token}", context.RemoteAddress(), acctok.Endpoint, acctok.FileName, context.TokenForLog());
        // note that here we exploit the non-dependency from fileid
        return await core.GetFile(context, acctok);
    }

    /// <summary>
    /// Returns a list of all currently open files, for operators only.
    /// This call is protected by the same shared secret as the /wopi/iop/openinapp call.
    /// </summary>
    IResult GetOpenFiles(HttpContext context)
    {
        if (!IsIopAuthorized(context))
        {
            log.LogWarning("iopGetOpenFiles: unauthorized access attempt, missing authorization token {client}", context.RemoteAddress());
            return Unauthorized();
        }

        // dump the current list of opened files in JSON format
        log.LogInformation("iopGetOpenFiles: returning list of open files {client}", context.RemoteAddress());
        return WopiResult.Json(state.OpenFilesJson());
    }

    /// <summary>
    /// Returns a list of all currently outstanding and resolved conflicted sessions, for operators only.
    /// This call is protected by the same shared secret as the /wopi/iop/openinapp call.
    /// </summary>
    IResult GetConflicts(HttpContext context)
    {
        if (!IsIopAuthorized(context))
        {
            log.LogWarning("iopGetConflicts: unauthorized access attempt, missing authorization token {client}", context.RemoteAddress());
            return Unauthorized();
        }

        // dump the current sets in JSON format
        log.LogInformation("iopGetConflicts: returning outstanding/resolved conflicted sessions {client}", context.RemoteAddress());
        return WopiResult.Json(state.ConflictsJson());
    }

    /// <summary>
    /// Returns a WOPI_URL and a WOPI_TOKEN values suitable as input for the WOPI validator test suite.
    /// This call is protected by the same shared secret as the /wopi/iop/openinapp call.
    /// Request arguments:
    /// - string filepath: the full path to the file used for the test. The file must exist
    /// - string endpoint (optional): the storage endpoint, defaults to 'default'
    /// - string usertoken: the credentials to access the given file (uid:gid for xrootd, bearer token for Reva)
    /// </summary>
    async Task<IResult> WopiTest(HttpContext context)
    {
        var request = context.Request;
        if (!IsIopAuthorized(context))
        {
            log.LogWarning("iopWopiTest: unauthorized access attempt, missing authorization token {client}", context.RemoteAddress());
            return Unauthorized();
        }

        // the Microsoft WOPI validator test suite requires to issue an access token for a predefined test file
        var filePath = request.Arg("filepath") ?? "";
        var endpoint = request.Arg("endpoint") ?? "default";
        var userToken = request.Arg("usertoken") ?? "";
        if (filePath.Length == 0 ||
            userToken.Length == 0)
        {
            return WopiResult.Text("Missing arguments", StatusCodes.Status400BadRequest);
        }

        if (state.UseHttps)
        {
            return WopiResult.Text("WOPI validator not supported in https mode", StatusCodes.Status400BadRequest);
        }

        var generated = await utils.GenerateAccessToken(
            userToken,
            filePath,
            ViewMode.ReadWrite,
            "test",
            $"test!{userToken}",
            UserType.Regular,
            "http://folderurlfortestonly/",
            endpoint,
            "WOPI validator",
            "http://fortestonly/",
            "http://fortestonly/",
            "TestTrace");
        log.LogInformation("iopWopiTest: preparing test via WOPI validator {client}", context.RemoteAddress());
        return WopiResult.Text($"-e WOPI_URL=http://localhost:{state.Port}/wopi/files/{generated.Inode} -e WOPI_TOKEN={generated.AccessToken}");
    }

    // WOPI protocol implementation

    /// <summary>
    /// The CheckFileInfo WOPI call.
    /// </summary>
    async Task<IResult> CheckFileInfo(HttpContext context, string fileid)
    {
        var (acctok, error) = await utils.ValidateAndLogHeaders(context, "CheckFileInfo");
        if (error != null)
        {
            return error;
        }

        return await core.CheckFileInfo(context, fileid, acctok!);
    }

    /// <summary>
    /// The GetFile WOPI call.
    /// </summary>
    async Task<IResult> GetFile(HttpContext context, string fileid)
    {
        var (acctok, error) = await utils.ValidateAndLogHeaders(context, "GetFile");
        if (error != null)
        {
            return error;
        }

        return await core.GetFile(context, acctok!);
    }

    /// <summary>
    /// A dispatcher method for all POST operations on files.
    /// </summary>
    async Task<IResult> FilesPost(HttpContext context, string fileid)
    {
        var request = context.Request;
        // must be one of the following strings
        var operation = request.Header("X-WOPI-Override");
        if (operation == null)
        {
            log.LogWarning("Missing argument {client} {requestedUrl} {error} {token}", context.Client(), request.BaseUrl(), "'X-WOPI-Override'", context.TokenForLog());
            return WopiResult.Text("Missing argument", StatusCodes.Status400BadRequest);
        }

        var (acctok, error) = await utils.ValidateAndLogHeaders(context, operation);
        if (error != null)
        {
            return error;
        }

        switch (operation)
        {
            case "GET_LOCK":
                return await core.GetLock(context, fileid, acctok!);
            case "PUT_USER_INFO":
                return await core.PutUserInfo(context, fileid, acctok!);
            case "PUT_RELATIVE":
                request.EnableBuffering();
                return await core.PutRelative(context, fileid, acctok!);
        }

        if (!acctok!.ViewMode.CanWrite())
        {
            // the remaining operations require write privileges
            return WopiResult.Text("Attempting to perform a write operation using a read-only token", StatusCodes.Status401Unauthorized);
        }

        switch (operation)
        {
            case "LOCK":
            case "REFRESH_LOCK":
                return await core.SetLock(context, fileid, acctok);
            case "UNLOCK":
                return await core.Unlock(context, fileid, acctok);
            case "DELETE":
                return await core.DeleteFile(context, fileid, acctok);
            case "RENAME_FILE":
                return await core.RenameFile(context, fileid, acctok);
        }

        // Any other op is unsupported
        log.LogWarning("Unknown/unsupported operation {operation}", operation);
        return WopiResult.Text("Not supported operation found in header", StatusCodes.Status501NotImplemented);
    }

    /// <summary>
    /// The PutFile WOPI call.
    /// </summary>
    async Task<IResult> PutFile(HttpContext context, string fileid)
    {
        var (acctok, error) = await utils.ValidateAndLogHeaders(context, "PutFile");
        if (error != null)
        {
            return error;
        }

        if (!acctok!.ViewMode.CanWrite())
        {
            // unlike the Python implementation, which only checked this for the POST operations on the file
            return WopiResult.Text("Attempting to perform a write operation using a read-only token", StatusCodes.Status401Unauthorized);
        }

        // buffered, so that the content can be stored for recovery after a failed write
        context.Request.EnableBuffering();
        return await core.PutFile(context, fileid, acctok!);
    }

    // Bridge functionality

    /// <summary>
    /// The WOPI bridge endpoint.
    /// </summary>
    async Task<IResult> BridgeRoot(HttpContext context, string docid)
    {
        if (!HttpMethods.IsHead(context.Request.Method))
        {
            // this is used via web hooks by the apps to save the document
            return await bridge.AppSave(context, docid);
        }

        // this is used to validate the access token
        var acctok = DecodeToken(context, out var error);
        var wopiSrc = context.Request.Arg("WOPISrc");
        if (acctok == null ||
            wopiSrc == null)
        {
            log.LogInformation("Bridge: expired or malformed token {client} {requestedUrl} {error} {token}", context.RemoteAddress(), context.Request.BaseUrl(), acctok == null ? error : "'WOPISrc'", acctok == null ? context.Request.Arg("access_token") ?? "N/A" : context.TokenForLog());
            return WopiResult.Text("Missing or invalid access token", StatusCodes.Status401Unauthorized);
        }

        if (bridge.ValidateDocId(wopiSrc, docid))
        {
            log.LogInformation("Bridge: access token is valid {filename} {token}", acctok.FileName, context.TokenForLog());
            return WopiResult.Text("");
        }

        log.LogInformation("Bridge: invalid docid for token {docid} {filename} {token}", docid, acctok.FileName, context.TokenForLog());
        return WopiResult.Text("Invalid docID for access token", StatusCodes.Status401Unauthorized);
    }
}
