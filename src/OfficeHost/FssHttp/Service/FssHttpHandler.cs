/// <summary>
/// The site of a file for desktop Office, at <c>{DesktopPath}/{token}</c>: the token of <see cref="IOfficeTokens"/> in the
/// path stands for the file and the user, as desktop Office sends no credentials of its own that a host can rely on
/// (Basic and forms-based sign-in are blocked by default: docs/fsshttp-notes.md). It answers what Word asks of a
/// SharePoint site that has the document:
/// - OPTIONS, with the X-MSFSSHTTP header that makes Word use the cell storage service, and the X-MSWEBURL of the site;
/// - the cell storage service ([MS-FSSHTTP]) at <c>_vti_bin/cellstorage.svc/CellStorageService</c>;
/// - IsOnlyClient of the shared access service ([MS-SHDACCWS]) at <c>_vti_bin/sharedaccess.asmx</c>, which Word polls
///   while alone;
/// - GET and HEAD of the document.
/// </summary>
class FssHttpHandler(
    IOfficeTokens tokens,
    IOfficeStorage storage,
    CellStorageService service,
    CoauthSessions sessions,
    FssHttpTargets targets,
    IOptions<OfficeHostOptions> options)
{
    const string defaultServerVersion = "16.0.0.0";
    const string methods = "GET, POST, OPTIONS, HEAD";

    public async Task Handle(HttpContext context, string token, string? rest)
    {
        var request = context.Request;
        var response = context.Response;
        var cancel = context.RequestAborted;
        var validated = await tokens.Validate(token, cancel);
        if (validated == null)
        {
            response.StatusCode = StatusCodes.Status401Unauthorized;
            return;
        }

        var settings = options.Value;
        var sitePath = $"{request.PathBase}{settings.DesktopPath.TrimEnd('/')}/{token}";
        var site = new FssHttpSite(validated.Access, $"{request.DesktopBaseUrl(settings)}{settings.DesktopPath.TrimEnd('/')}/{token}", token, validated.Expires);
        response.Headers["MicrosoftSharePointTeamServices"] = settings.ServerVersion ?? defaultServerVersion;
        var path = rest ?? "";
        var method = request.Method;
        if (HttpMethods.IsOptions(method))
        {
            Options(request, response, sitePath);
            return;
        }

        if (HttpMethods.IsPost(method))
        {
            if (path.EndsWith("_vti_bin/cellstorage.svc/CellStorageService", StringComparison.OrdinalIgnoreCase))
            {
                await CellStorage(context, site, cancel);
                return;
            }

            if (path.EndsWith("_vti_bin/sharedaccess.asmx", StringComparison.OrdinalIgnoreCase))
            {
                await IsOnlyClient(context, site, cancel);
                return;
            }

            response.StatusCode = StatusCodes.Status404NotFound;
            return;
        }

        if ((HttpMethods.IsGet(method) || HttpMethods.IsHead(method)) &&
            !path.Contains("_vti_", StringComparison.OrdinalIgnoreCase))
        {
            await Document(context, site, path, cancel);
            return;
        }

        response.StatusCode = StatusCodes.Status404NotFound;
    }

    /// <summary>
    /// The OPTIONS answer of SharePoint once signed in (docs/fsshttp-notes.md): X-MSFSSHTTP, which no published spec has,
    /// makes Word use the cell storage service, and X-MSWEBURL, sent when Word asks with X-MSGETWEBURL, is the path of
    /// the site the service is under. The WebDAV headers of [MS-WDVME] 3.2.5.1 go with them.
    /// </summary>
    static void Options(HttpRequest request, HttpResponse response, string sitePath)
    {
        var headers = response.Headers;
        headers["X-MSFSSHTTP"] = "1.3";
        if (request.Headers["X-MSGETWEBURL"] == "t")
        {
            headers["X-MSWEBURL"] = sitePath;
        }

        headers.Allow = methods;
        headers["Public"] = methods;
        headers["DAV"] = "1,2";
        headers["MS-Author-Via"] = "DAV";
        headers["DocumentManagementServer"] = "Properties Schema;Source Control;Version History;";
        headers.AcceptRanges = "none";
        response.StatusCode = StatusCodes.Status200OK;
    }

    async Task CellStorage(HttpContext context, FssHttpSite site, Cancel cancel)
    {
        // the changes of a file are as large as the file, which a writer can send up to MaxFileSize, as through WOPI
        var limit = context.Features.Get<IHttpMaxRequestBodySizeFeature>();
        if (site.Access.CanWrite &&
            limit is { IsReadOnly: false })
        {
            limit.MaxRequestBodySize = options.Value.MaxFileSize;
        }

        SoapMessage message;
        try
        {
            message = await SoapMessage.Read(context.Request, cancel);
        }
        catch (InvalidDataException)
        {
            context.Response.StatusCode = StatusCodes.Status400BadRequest;
            return;
        }
        catch (BadHttpRequestException exception)
        {
            context.Response.StatusCode = exception.StatusCode;
            return;
        }

        var mtom = new MtomResponse();
        XElement envelope;
        try
        {
            envelope = await service.Execute(message, site, mtom, cancel);
        }
        catch (InvalidDataException)
        {
            context.Response.StatusCode = StatusCodes.Status400BadRequest;
            return;
        }

        await mtom.Write(context.Response, envelope, cancel);
    }

    /// <summary>
    /// IsOnlyClient ([MS-SHDACCWS]): false while a coauthoring transition is pending or others share the lock, so that a
    /// client alone starts polling the coauthoring status ([MS-FSSHTTP] 3.1.4.3.6). Word sends it only while alone, every
    /// 8 to 30 seconds and with no client id, so it keeps the clients of the session in, as the coauthoring status poll
    /// does once there are several.
    /// </summary>
    async Task IsOnlyClient(HttpContext context, FssHttpSite site, Cancel cancel)
    {
        // the id is the ResourceID of the document: after a Save As, that of the new file
        string? id = null;
        try
        {
            var document = await XDocument.LoadAsync(context.Request.Body, LoadOptions.None, cancel);
            id = document.Descendants().FirstOrDefault(_ => _.Name.LocalName == "id")?.Value;
        }
        catch (XmlException)
        {
        }

        if (Guid.TryParse(id, out var resourceId))
        {
            site = await targets.Find(site, resourceId, cancel) ?? site;
        }

        // only a writer, who can be in the session, keeps it alive: a reader would keep the lock of a writer gone
        Session session;
        if (site.Access.CanWrite)
        {
            var result = await sessions.Change(
                site.FileId,
                _ => new("Success", _ with
                {
                    Clients = _.Clients.Select(client => client with { Expires = sessions.Now.AddSeconds(client.TimeoutSeconds) }).ToList()
                }),
                cancel);
            session = result.Session;
        }
        else
        {
            session = await sessions.Get(site.FileId, cancel);
        }

        var alone = session.Clients.Count <= 1 && !session.TransitionPending;
        var soap = SoapMessage.Soap;
        var envelope = new XElement(
            SoapMessage.Envelope + "Envelope",
            new XAttribute(XNamespace.Xmlns + "soap", SoapMessage.Envelope),
            new XElement(
                SoapMessage.Envelope + "Body",
                new XElement(
                    soap + "IsOnlyClientResponse",
                    new XAttribute("xmlns", soap.NamespaceName),
                    new XElement(soap + "IsOnlyClientResult", XmlConvert.ToString(alone)))));
        context.Response.ContentType = "text/xml; charset=utf-8";
        await context.Response.WriteAsync(new XDeclaration("1.0", "utf-8", null) + envelope.ToString(SaveOptions.DisableFormatting), cancel);
    }

    /// <summary>
    /// The document the path names (<see cref="FssHttpTargets"/>), for a client that falls back from the cell storage
    /// service. Before a Save As, Word checks with a HEAD that the new name is free: a file of that name that the site
    /// could take over is answered without its content.
    /// </summary>
    async Task Document(HttpContext context, FssHttpSite site, string path, Cancel cancel)
    {
        var response = context.Response;
        string? name = null;
        if (path.Length > 0)
        {
            name = path;
        }

        var target = await targets.Find(site, name, cancel);
        OfficeFile? file;
        if (target.Site != null)
        {
            file = await storage.GetFile(target.Site.FileId, cancel);
        }
        else if (target.SaveAsExisting != null &&
                 HttpMethods.IsHead(context.Request.Method))
        {
            file = target.SaveAsExisting;
        }
        else
        {
            file = null;
        }

        if (file == null)
        {
            response.StatusCode = StatusCodes.Status404NotFound;
            return;
        }

        response.ContentType = "application/octet-stream";
        response.ContentLength = file.Size;
        response.Headers.ETag = $"\"{file.Version}\"";
        response.Headers.LastModified = file.LastModified.ToString("R");
        if (HttpMethods.IsHead(context.Request.Method))
        {
            return;
        }

        await using var stream = await storage.OpenRead(file.Id, cancel);
        await stream.CopyToAsync(response.Body, cancel);
    }
}
