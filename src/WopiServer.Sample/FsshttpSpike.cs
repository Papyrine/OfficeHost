using System.Text;
using System.Xml.Linq;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Negotiate;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;

/// <summary>
/// The throwaway spike of the FSSHTTP Phase 0 (docs/fsshttp-spike.md): just enough of the SharePoint discovery sequence
/// and of [MS-FSSHTTP] for desktop Word to show what it asks of a host that is not SharePoint. Every request and response
/// is logged in full to the spike-logs folder. Not part of the library: enabled by the Spike:Fsshttp setting only.
/// </summary>
/// <remarks>
/// A site is mounted five times: /t/{token} (the WOPI token in the path), /o/{token} (the token in the path, behind the
/// Office Forms Based Authentication of [MS-OFBA]), /b (Basic auth over HTTPS, any user name, a token as the password),
/// /n (Windows auth) and / (anonymous, to see the probes Word sends to the root).
/// </remarks>
static class FsshttpSpike
{
    static readonly XNamespace soap = "http://schemas.microsoft.com/sharepoint/soap/";
    static readonly XNamespace envelope = "http://schemas.xmlsoap.org/soap/envelope/";
    static readonly XNamespace dav = "DAV:";
    const string serverVersion = "16.0.0.0";
    const string docxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    // E_FAIL, as SharePoint sends with an error code (MS-FSSHTTP 4.1.2)
    const long failHResult = 2147500037;
    static int sequence;

    /// <summary>
    /// The Windows auth (Negotiate: Kerberos or NTLM) of the /n site, when Spike:Negotiate is true: it needs Kestrel,
    /// so it is off in the tests, on their test server.
    /// </summary>
    public static void AddFsshttpSpike(this IServiceCollection services, IConfiguration configuration)
    {
        if (configuration.GetValue<bool>("Spike:Negotiate"))
        {
            services.AddAuthentication(NegotiateDefaults.AuthenticationScheme)
                .AddNegotiate();
        }
    }

    public static void UseFsshttpSpike(this WebApplication app, FileSystemStorage storage)
    {
        var logs = Path.Combine(app.Environment.ContentRootPath, app.Configuration["Spike:Logs"] ?? "spike-logs");
        Directory.CreateDirectory(logs);
        var tokens = app.Services.GetRequiredService<IWopiTokens>();
        var logger = app.Logger;
        app.Use((context, next) => Log(context, next, logs, logger));
        // headers to try on OPTIONS without a rebuild, as those of a SharePoint trace: Spike:OptionsHeaders:{name}
        var optionsHeaders = app.Configuration.GetSection("Spike:OptionsHeaders")
            .GetChildren()
            .ToDictionary(_ => _.Key, _ => _.Value ?? "");
        var negotiate = app.Configuration.GetValue<bool>("Spike:Negotiate");
        app.Use((context, next) => Handle(context, next, storage, tokens, optionsHeaders, negotiate));
        app.MapGet("/spike", (HttpRequest request) => Launch(request, storage, tokens));
    }

    /// <summary>
    /// The page with the ms-word links (MS-OFFICE URI schemes: ofe = open for edit, u = the URL) of the welcome document.
    /// </summary>
    static async Task<IResult> Launch(HttpRequest request, FileSystemStorage storage, IWopiTokens tokens)
    {
        var fileId = storage.GetFileId("Welcome.docx");
        var token = await tokens.Issue(Access(fileId, "demo", "Demo User"));
        var baseUrl = BaseUrl(request);
        var links = new[]
        {
            ("Token in the path", $"ms-word:ofe|u|{baseUrl}/t/{token.Value}/Welcome.docx"),
            ("Forms based auth (MS-OFBA), signing in from the token in the path", $"ms-word:ofe|u|{baseUrl}/o/{token.Value}/Welcome.docx"),
            ("Basic auth (any user name, the token below as the password)", $"ms-word:ofe|u|{baseUrl}/b/Welcome.docx"),
            ("Windows auth (Negotiate)", $"ms-word:ofe|u|{baseUrl}/n/Welcome.docx"),
            ("Anonymous, at the root", $"ms-word:ofe|u|{baseUrl}/Welcome.docx")
        };
        var html = new StringBuilder("<!doctype html><title>FSSHTTP spike</title><h1>FSSHTTP spike</h1><ul>");
        foreach (var (label, link) in links)
        {
            html.Append($"<li>{WebUtility.HtmlEncode(label)}: <a href=\"{WebUtility.HtmlEncode(link)}\">{WebUtility.HtmlEncode(link)}</a></li>");
        }

        html.Append($"</ul><p>Token: <code>{WebUtility.HtmlEncode(token.Value)}</code></p>");
        return Results.Content(html.ToString(), "text/html");
    }

    static WopiAccess Access(string fileId, string userId, string userName) =>
        new(fileId, userId)
        {
            UserName = userName,
            CanWrite = true
        };

    static string BaseUrl(HttpRequest request) => $"{request.Scheme}://{request.Host}{request.PathBase}";

    #region Logging

    static async Task Log(HttpContext context, Func<Task> next, string folder, ILogger logger)
    {
        var number = Interlocked.Increment(ref sequence);
        var request = context.Request;
        request.EnableBuffering();
        var requestBody = await ReadAll(request.Body);
        request.Body.Position = 0;

        var response = context.Response;
        var original = response.Body;
        using var captured = new MemoryStream();
        response.Body = captured;
        try
        {
            await next();
        }
        finally
        {
            response.Body = original;
            var responseBody = captured.ToArray();
            var log = new StringBuilder();
            log.AppendLine($"#{number} {DateTimeOffset.Now:O}");
            log.AppendLine($"{request.Method} {request.Scheme}://{request.Host}{request.PathBase}{request.Path}{request.QueryString} {request.Protocol}");
            AppendHeaders(log, request.Headers);
            log.AppendLine();
            await AppendBody(log, requestBody, request.ContentType);
            log.AppendLine();
            log.AppendLine($"--- response {response.StatusCode}");
            AppendHeaders(log, response.Headers);
            log.AppendLine();
            await AppendBody(log, responseBody, response.ContentType);
            var name = $"{number:0000}-{request.Method}-{FileName(request.Path.Value)}.txt";
            await File.WriteAllTextAsync(Path.Combine(folder, name), log.ToString());
            logger.LogInformation(
                "Spike #{Number} {Method} {Path} {UserAgent} -> {Status}",
                number,
                request.Method,
                request.Path.Value,
                request.Headers.UserAgent.ToString(),
                response.StatusCode);
            if (responseBody.Length > 0 && !HttpMethods.IsHead(request.Method))
            {
                await original.WriteAsync(responseBody);
            }
        }
    }

    static async Task<byte[]> ReadAll(Stream stream)
    {
        using var buffer = new MemoryStream();
        await stream.CopyToAsync(buffer);
        return buffer.ToArray();
    }

    static void AppendHeaders(StringBuilder log, IHeaderDictionary headers)
    {
        foreach (var (key, value) in headers)
        {
            log.AppendLine($"{key}: {value}");
        }
    }

    static async Task AppendBody(StringBuilder log, byte[] body, string? contentType)
    {
        if (body.Length == 0)
        {
            log.AppendLine("(no body)");
            return;
        }

        if (IsText(contentType))
        {
            log.AppendLine(Encoding.UTF8.GetString(body));
            return;
        }

        AppendHex(log, body);
        if (contentType != null &&
            contentType.Contains("multipart/", StringComparison.OrdinalIgnoreCase))
        {
            await AppendParts(log, body, contentType);
        }
    }

    static bool IsText(string? contentType)
    {
        if (contentType == null ||
            contentType.Contains("multipart/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string[] texts = ["text/", "xml", "json", "urlencoded", "vermeer", "html"];
        return texts.Any(_ => contentType.Contains(_, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// A hex dump: offset, 16 bytes, and their printable ASCII.
    /// </summary>
    static void AppendHex(StringBuilder log, byte[] body)
    {
        log.AppendLine($"({body.Length} bytes, hex)");
        for (var offset = 0; offset < body.Length; offset += 16)
        {
            var line = body.AsSpan(offset, Math.Min(16, body.Length - offset));
            log.Append($"{offset:x8}  ");
            for (var index = 0; index < 16; index++)
            {
                if (index < line.Length)
                {
                    log.Append($"{line[index]:x2} ");
                }
                else
                {
                    log.Append("   ");
                }
            }

            log.Append(' ');
            foreach (var value in line)
            {
                if (value is >= 0x20 and < 0x7f)
                {
                    log.Append((char) value);
                }
                else
                {
                    log.Append('.');
                }
            }

            log.AppendLine();
        }
    }

    /// <summary>
    /// The parts of a multipart body (MTOM/XOP), each as text or hex.
    /// </summary>
    static async Task AppendParts(StringBuilder log, byte[] body, string contentType)
    {
        try
        {
            var reader = new MultipartReader(Boundary(contentType), new MemoryStream(body));
            var index = 0;
            while (await reader.ReadNextSectionAsync() is { } section)
            {
                log.AppendLine($"--- part {index++}");
                foreach (var (key, value) in section.Headers ?? [])
                {
                    log.AppendLine($"{key}: {value}");
                }

                log.AppendLine();
                var content = await ReadAll(section.Body);
                if (section.ContentType != null && section.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase))
                {
                    log.AppendLine(Encoding.UTF8.GetString(content));
                }
                else
                {
                    AppendHex(log, content);
                }
            }
        }
        catch (Exception exception)
        {
            log.AppendLine($"(could not read the parts: {exception.Message})");
        }
    }

    static string Boundary(string contentType) =>
        HeaderUtilities.RemoveQuotes(MediaTypeHeaderValue.Parse(contentType).Boundary).Value!;

    static string FileName(string? path)
    {
        var name = path ?? "";
        if (name.StartsWith("/t/", StringComparison.Ordinal) ||
            name.StartsWith("/o/", StringComparison.Ordinal))
        {
            // the token, which is long
            var end = name.IndexOf('/', 3);
                name = name[..3] + "~" + name[(end < 0 ? name.Length : end)..];
        }

        var builder = new StringBuilder();
        foreach (var value in name.Trim('/'))
        {
            if (char.IsLetterOrDigit(value) || value is '.' or '-' or '_')
            {
                builder.Append(value);
            }
            else
            {
                builder.Append('_');
            }
        }

        if (builder.Length > 80)
        {
            builder.Length = 80;
        }

        return builder.ToString();
    }

    #endregion

    #region Dispatch

    /// <summary>
    /// The site a request is for, the path in it, and the access its credentials give.
    /// </summary>
    record Site(string Path, string Rest, string? UserId, string? UserName);

    static async Task Handle(
        HttpContext context,
        Func<Task> next,
        FileSystemStorage storage,
        IWopiTokens tokens,
        Dictionary<string, string> optionsHeaders,
        bool negotiate)
    {
        var path = context.Request.Path.Value ?? "/";
        var method = context.Request.Method;
        if (path == "/spike")
        {
            await next();
            return;
        }

        var response = context.Response;
        Site site;
        if (path.StartsWith("/t/", StringComparison.Ordinal))
        {
            var end = path.IndexOf('/', 3);
            if (end < 0)
            {
                end = path.Length;
            }

            var token = await tokens.Validate(path[3..end]);
            if (token == null)
            {
                response.StatusCode = StatusCodes.Status401Unauthorized;
                return;
            }

            site = new(path[..end], path[end..], token.Access.UserId, token.Access.UserName);
        }
        else if (path.StartsWith("/o/", StringComparison.Ordinal))
        {
            var end = path.IndexOf('/', 3);
            if (end < 0)
            {
                end = path.Length;
            }

            var sitePath = path[..end];
            var siteToken = path[3..end];
            var token = await tokens.Validate(siteToken);
            if (token == null)
            {
                response.StatusCode = StatusCodes.Status404NotFound;
                return;
            }

            if (await FormsLogin(context, sitePath, siteToken, token.Access.UserId))
            {
                return;
            }

            site = new(sitePath, path[end..], token.Access.UserId, token.Access.UserName);
        }
        else if (negotiate &&
                 (path == "/n" ||
                  path.StartsWith("/n/", StringComparison.Ordinal)))
        {
            // Windows auth: Office does not class it as an insecure sign-in method, as it does Basic and MS-OFBA
            var result = await context.AuthenticateAsync(NegotiateDefaults.AuthenticationScheme);
            if (!result.Succeeded)
            {
                await context.ChallengeAsync(NegotiateDefaults.AuthenticationScheme);
                return;
            }

            var name = result.Principal.Identity?.Name ?? "windows";
            site = new("/n", path[2..], name, name);
        }
        else if (path == "/b" ||
                 path.StartsWith("/b/", StringComparison.Ordinal))
        {
            var user = await BasicUser(context.Request, tokens);
            if (user == null)
            {
                // MS-FSSHTTP 2.1: the transport is HTTPS, with the authentication of the server
                response.StatusCode = StatusCodes.Status401Unauthorized;
                response.Headers.WWWAuthenticate = "Basic realm=\"WopiServer spike\"";
                return;
            }

            site = new("/b", path[2..], user, user);
        }
        else
        {
            site = new("", path, null, null);
        }

        var rest = site.Rest;
        if (!IsSpike(method, rest, site))
        {
            await next();
            return;
        }

        response.Headers["MicrosoftSharePointTeamServices"] = serverVersion;
        response.Headers["X-SharePointHealthScore"] = "0";
        if (HttpMethods.IsOptions(method))
        {
            Options(context.Request, response, site);
            foreach (var (name, value) in optionsHeaders)
            {
                response.Headers[name] = value;
            }

            return;
        }

        if (rest.EndsWith("/_vti_inf.html", StringComparison.OrdinalIgnoreCase))
        {
            await VtiInf(response);
            return;
        }

        if (Contains(rest, "/_vti_bin/shtml.dll") ||
            Contains(rest, "/_vti_bin/_vti_aut/author.dll") ||
            Contains(rest, "/_vti_bin/_vti_adm/admin.dll"))
        {
            await Fpse(context, site);
            return;
        }

        if (Contains(rest, "/_vti_bin/webs.asmx"))
        {
            await Webs(context, site);
            return;
        }

        if (Contains(rest, "/_vti_bin/cellstorage.svc"))
        {
            await CellStorage(context, site, storage);
            return;
        }

        if (Contains(rest, "/_vti_bin/") ||
            Contains(rest, "/_api/") ||
            Contains(rest, "/_layouts/"))
        {
            response.StatusCode = StatusCodes.Status404NotFound;
            return;
        }

        switch (method)
        {
            case "PROPFIND":
                await Propfind(context, site, storage);
                return;
            case "LOCK":
                await Lock(response);
                return;
            case "UNLOCK":
                response.StatusCode = StatusCodes.Status204NoContent;
                return;
        }

        await Document(context, site, storage);
    }

    /// <summary>
    /// Whether the spike answers the request, rather than the sample: everything in /t and /b, and at the root the
    /// probes of Office (OPTIONS, the FrontPage and SharePoint paths, WebDAV methods) and the documents.
    /// </summary>
    static bool IsSpike(string method, string rest, Site site)
    {
        if (site.Path.Length > 0)
        {
            return true;
        }

        if (HttpMethods.IsOptions(method) ||
            method is "PROPFIND" or "LOCK" or "UNLOCK" ||
            Contains(rest, "/_vti_") ||
            Contains(rest, "/_api/") ||
            Contains(rest, "/_layouts/"))
        {
            return true;
        }

        if (rest.StartsWith("/wopi", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return rest.EndsWith(".docx", StringComparison.OrdinalIgnoreCase);
    }

    static bool Contains(string path, string value) =>
        path.Contains(value, StringComparison.OrdinalIgnoreCase);

    const string formsCookie = "spike-fba";

    /// <summary>
    /// Office Forms Based Authentication ([MS-OFBA] 2.2.1, 2.2.2) at /o/{token}: without the cookie, a 403 with the
    /// login and return pages; the login page signs in from the token of the site path, with nothing to type, and
    /// redirects to the return page that ends the login dialog of Office. Returns whether the request was answered.
    /// </summary>
    static async Task<bool> FormsLogin(HttpContext context, string sitePath, string siteToken, string userId)
    {
        var request = context.Request;
        var response = context.Response;
        var rest = request.Path.Value![sitePath.Length..];
        var login = $"{BaseUrl(request)}{sitePath}/_login";
        var done = $"{login}/done/";
        if (rest == "/_login")
        {
            response.Cookies.Append(
                formsCookie,
                siteToken,
                new()
                {
                    Path = sitePath,
                    Secure = true,
                    HttpOnly = true,
                    SameSite = Microsoft.AspNetCore.Http.SameSiteMode.None
                });
            response.Redirect(done);
            return true;
        }

        if (rest == "/_login/done/")
        {
            response.ContentType = "text/html";
            await response.WriteAsync($"<!doctype html><title>Signed in</title><p>Signed in as {WebUtility.HtmlEncode(userId)}</p>");
            return true;
        }

        if (request.Cookies[formsCookie] == siteToken)
        {
            return false;
        }

        response.StatusCode = StatusCodes.Status403Forbidden;
        response.Headers["X-FORMS_BASED_AUTH_REQUIRED"] = login;
        response.Headers["X-FORMS_BASED_AUTH_RETURN_URL"] = done;
        response.Headers["X-FORMS_BASED_AUTH_DIALOG_SIZE"] = "800x600";
        return true;
    }

    static async Task<string?> BasicUser(HttpRequest request, IWopiTokens tokens)
    {
        var header = request.Headers.Authorization.ToString();
        if (!header.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        string credentials;
        try
        {
            credentials = Encoding.UTF8.GetString(Convert.FromBase64String(header[6..].Trim()));
        }
        catch (FormatException)
        {
            return null;
        }

        var colon = credentials.IndexOf(':');
        if (colon <= 0)
        {
            return null;
        }

        if (await tokens.Validate(credentials[(colon + 1)..]) == null)
        {
            return null;
        }

        return credentials[..colon];
    }

    static async Task<WopiFile?> FindFile(FileSystemStorage storage, string rest)
    {
        var relative = Uri.UnescapeDataString(rest.TrimStart('/'));
        if (relative.Length == 0)
        {
            return null;
        }

        try
        {
            var fileId = storage.GetFileId(relative);
            return await storage.GetFile(fileId);
        }
        catch (Exception exception) when (exception is FileNotFoundException or UnauthorizedAccessException or DirectoryNotFoundException)
        {
            return null;
        }
    }

    #endregion

    #region WebDAV and FrontPage

    /// <summary>
    /// The OPTIONS answer of SharePoint: the WebDAV methods with GETLIB, MS-Author-Via, DocumentManagementServer and
    /// Public-Extension ([MS-WDVME] 3.2.5.1, 4.1.2), and X-MSDAVEXT ([MS-WEBDAVE] 2.2.1). Then two headers that no
    /// published spec has, from a capture of SharePoint Server 16.0.0.5522 (docs/fsshttp-notes.md): X-MSFSSHTTP, which
    /// SharePoint sends on every OPTIONS once signed in, and X-MSWEBURL, the path of the site, which it sends when Word
    /// asks with X-MSGETWEBURL: t.
    /// </summary>
    static void Options(HttpRequest request, HttpResponse response, Site site)
    {
        const string methods = "GET, POST, OPTIONS, HEAD, MKCOL, PUT, PROPFIND, PROPPATCH, DELETE, MOVE, COPY, GETLIB, LOCK, UNLOCK";
        var headers = response.Headers;
        headers["X-MSFSSHTTP"] = "1.3";
        if (request.Headers["X-MSGETWEBURL"] == "t")
        {
            headers["X-MSWEBURL"] = site.Path.Length == 0 ? "/" : site.Path;
        }

        headers.Allow = methods;
        headers["Public"] = methods;
        headers["DAV"] = "1,2";
        headers["MS-Author-Via"] = "MS-FP/4.0,DAV";
        headers["DocumentManagementServer"] = "Properties Schema;Source Control;Version History;";
        headers["Public-Extension"] = "http://schemas.microsoft.com/repl-2";
        headers["X-MSDAVEXT"] = "1";
        headers.AcceptRanges = "none";
        response.StatusCode = StatusCodes.Status200OK;
    }

    /// <summary>
    /// The FrontPage configuration of the server ([MS-FPSE] 3.1.5.1, _vti_inf.html).
    /// </summary>
    static Task VtiInf(HttpResponse response)
    {
        response.ContentType = "text/html";
        return response.WriteAsync(
            $"""
             <!-- FrontPage Configuration Information
             FPVersion="{serverVersion}"
             FPShtmlScriptUrl="_vti_bin/shtml.dll/_vti_rpc"
             FPAuthorScriptUrl="_vti_bin/_vti_aut/author.dll"
             FPAdminScriptUrl="_vti_bin/_vti_adm/admin.dll"
             TPScriptUrl="_vti_bin/owssvr.dll"
             -->
             """);
    }

    /// <summary>
    /// The FrontPage RPC methods ([MS-FPSE] 3.1.4): server version, url to web url and open service; any other
    /// method fails, so that the log shows it.
    /// </summary>
    static async Task Fpse(HttpContext context, Site site)
    {
        var body = await new StreamReader(context.Request.Body).ReadToEndAsync();
        var form = QueryHelpers.ParseQuery(body);
        var method = form.TryGetValue("method", out var value) ? value.ToString() : "";
        var name = method.Split(':')[0];
        string content;
        switch (name)
        {
            case "server version":
                content =
                    $"""
                     <p>method={method}
                     <p>server version=
                     <ul>
                     <li>major ver=16
                     <li>minor ver=0
                     <li>phase ver=0
                     <li>ver incr=0
                     </ul>
                     <p>source control=1
                     """;
                break;
            case "url to web url":
                var url = form.TryGetValue("url", out var urlValue) ? urlValue.ToString() : "";
                var file = url.StartsWith(site.Path, StringComparison.Ordinal) ? url[site.Path.Length..] : url;
                content =
                    $"""
                     <p>method={method}
                     <p>webUrl={site.Path}
                     <p>fileUrl={file.TrimStart('/')}
                     """;
                break;
            case "open service":
                content =
                    $"""
                     <p>method={method}
                     <p>service=
                     <ul>
                     <li>service_name=
                     <li>meta_info=
                     <ul>
                     <li>vti_casesensitiveurls
                     <li>IX|0
                     <li>vti_longfilenames
                     <li>IX|1
                     <li>vti_welcomenames
                     <li>VX|
                     <li>vti_username
                     <li>SX|{site.UserName}
                     <li>vti_servertz
                     <li>SX|+0000
                     <li>vti_sourcecontrolsystem
                     <li>SR|lw
                     <li>vti_sourcecontrolversion
                     <li>SR|V1
                     <li>vti_doclibwebviewenabled
                     <li>IX|0
                     <li>vti_sourcecontrolcookie
                     <li>SX|fp_internal
                     <li>vti_sourcecontrolproject
                     <li>SX|&#60;STS-based Locking&#62;
                     </ul>
                     </ul>
                     """;
                break;
            default:
                content =
                    $"""
                     <p>method={method}
                     <p>status=
                     <ul>
                     <li>status=589826
                     <li>osstatus=0
                     <li>msg=The method is not supported by the spike
                     <li>osmsg=
                     </ul>
                     """;
                break;
        }

        context.Response.ContentType = "application/x-vermeer-rpc";
        await context.Response.WriteAsync(
            $"""
             <html><head><title>vermeer RPC packet</title></head>
             <body>
             {content}
             </body>
             </html>
             """);
    }

    /// <summary>
    /// The Webs web service ([MS-WEBSS] 3.1.4.x WebUrlFromPageUrl): the site of a page.
    /// </summary>
    static async Task Webs(HttpContext context, Site site)
    {
        var action = context.Request.Headers["SOAPAction"].ToString().Trim('"');
        var response = context.Response;
        response.ContentType = "text/xml; charset=utf-8";
        if (!action.EndsWith("/WebUrlFromPageUrl", StringComparison.Ordinal))
        {
            response.StatusCode = StatusCodes.Status500InternalServerError;
            await response.WriteAsync(SoapFault($"The spike does not support {action}"));
            return;
        }

        var xml = new XDocument(
            new XElement(
                envelope + "Envelope",
                new XAttribute(XNamespace.Xmlns + "soap", envelope),
                new XElement(
                    envelope + "Body",
                    new XElement(
                        soap + "WebUrlFromPageUrlResponse",
                        new XElement(soap + "WebUrlFromPageUrlResult", BaseUrl(context.Request) + site.Path)))));
        await response.WriteAsync(xml.ToString(SaveOptions.DisableFormatting));
    }

    static string SoapFault(string message) =>
        new XElement(
                envelope + "Envelope",
                new XAttribute(XNamespace.Xmlns + "soap", envelope),
                new XElement(
                    envelope + "Body",
                    new XElement(
                        envelope + "Fault",
                        new XElement("faultcode", "soap:Server"),
                        new XElement("faultstring", message))))
            .ToString(SaveOptions.DisableFormatting);

    /// <summary>
    /// The properties of the document, or of the site folder ([RFC4918] PROPFIND, [MS-WEBDAVE]).
    /// </summary>
    static async Task Propfind(HttpContext context, Site site, FileSystemStorage storage)
    {
        var file = await FindFile(storage, site.Rest);
        var href = BaseUrl(context.Request) + site.Path + site.Rest;
        XElement prop;
        if (file == null)
        {
            prop = new(
                dav + "prop",
                new XElement(dav + "displayname", Path.GetFileName(site.Rest.TrimEnd('/'))),
                new XElement(dav + "resourcetype", new XElement(dav + "collection")),
                new XElement(dav + "getlastmodified", DateTimeOffset.UtcNow.ToString("R")));
        }
        else
        {
            prop = new(
                dav + "prop",
                new XElement(dav + "displayname", file.Name),
                new XElement(dav + "resourcetype"),
                new XElement(dav + "getcontentlength", file.Size),
                new XElement(dav + "getcontenttype", docxType),
                new XElement(dav + "getetag", $"\"{file.Version}\""),
                new XElement(dav + "getlastmodified", file.LastModified.ToString("R")));
        }

        var xml = new XElement(
            dav + "multistatus",
            new XAttribute(XNamespace.Xmlns + "D", dav),
            new XElement(
                dav + "response",
                new XElement(dav + "href", href),
                new XElement(
                    dav + "propstat",
                    prop,
                    new XElement(dav + "status", "HTTP/1.1 200 OK"))));
        context.Response.StatusCode = StatusCodes.Status207MultiStatus;
        context.Response.ContentType = "text/xml; charset=utf-8";
        await context.Response.WriteAsync(xml.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// A fake exclusive WebDAV lock ([RFC4918] LOCK), so that a fallback of Word to WebDAV shows in the log.
    /// </summary>
    static async Task Lock(HttpResponse response)
    {
        var token = $"opaquelocktoken:{Guid.NewGuid()}";
        var xml = new XElement(
            dav + "prop",
            new XAttribute(XNamespace.Xmlns + "D", dav),
            new XElement(
                dav + "lockdiscovery",
                new XElement(
                    dav + "activelock",
                    new XElement(dav + "locktype", new XElement(dav + "write")),
                    new XElement(dav + "lockscope", new XElement(dav + "exclusive")),
                    new XElement(dav + "depth", "0"),
                    new XElement(dav + "timeout", "Second-3600"),
                    new XElement(dav + "locktoken", new XElement(dav + "href", token)))));
        response.Headers["Lock-Token"] = $"<{token}>";
        response.ContentType = "text/xml; charset=utf-8";
        await response.WriteAsync(xml.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>
    /// GET, HEAD and PUT of the document itself.
    /// </summary>
    static async Task Document(HttpContext context, Site site, FileSystemStorage storage)
    {
        var response = context.Response;
        var file = await FindFile(storage, site.Rest);
        if (file == null)
        {
            response.StatusCode = StatusCodes.Status404NotFound;
            return;
        }

        var method = context.Request.Method;
        if (HttpMethods.IsPut(method))
        {
            var written = await storage.Write(file.Id, context.Request.Body);
            response.Headers.ETag = $"\"{written.Version}\"";
            response.StatusCode = StatusCodes.Status204NoContent;
            return;
        }

        if (!HttpMethods.IsGet(method) &&
            !HttpMethods.IsHead(method))
        {
            response.StatusCode = StatusCodes.Status405MethodNotAllowed;
            return;
        }

        response.ContentType = docxType;
        response.ContentLength = file.Size;
        response.Headers.ETag = $"\"{file.Version}\"";
        response.Headers.LastModified = file.LastModified.ToString("R");
        if (HttpMethods.IsHead(method))
        {
            return;
        }

        await using var stream = await storage.OpenRead(file.Id);
        await stream.CopyToAsync(response.Body);
    }

    #endregion

    #region Cell storage

    /// <summary>
    /// The cell storage service ([MS-FSSHTTP] 2.2.2), at the site or at the document URL (1.5): a SOAP request of
    /// RequestVersion and RequestCollection, either a plain SOAP body or MTOM/XOP (2.1). The spike answers WhoAmI, ServerTime and GetDocMetaInfo, stubs the
    /// XML-only lock and editors subrequests with a success, and fails the Cell subrequests, which need FSSHTTPB.
    /// </summary>
    static async Task CellStorage(HttpContext context, Site site, FileSystemStorage storage)
    {
        var request = context.Request;
        var mtom = request.ContentType != null &&
                   request.ContentType.Contains("multipart/related", StringComparison.OrdinalIgnoreCase);
        XDocument xml;
        if (mtom)
        {
            var reader = new MultipartReader(Boundary(request.ContentType!), request.Body);
            var section = await reader.ReadNextSectionAsync();
            xml = await XDocument.LoadAsync(section!.Body, LoadOptions.None, context.RequestAborted);
        }
        else
        {
            xml = await XDocument.LoadAsync(request.Body, LoadOptions.None, context.RequestAborted);
        }

        var baseUrl = BaseUrl(request);
        var minorVersion = (string?) xml.Descendants(soap + "RequestVersion").FirstOrDefault()?.Attribute("MinorVersion") ?? "2";
        var responses = new List<XElement>();
        foreach (var element in xml.Descendants(soap + "Request"))
        {
            var url = (string?) element.Attribute("Url") ?? "";
            var file = await FileOfUrl(url, baseUrl, site, storage);
            responses.Add(Response(element, url, site, file));
        }
        var body = new XElement(
            envelope + "Envelope",
            new XAttribute(XNamespace.Xmlns + "s", envelope),
            new XElement(
                envelope + "Body",
                // MS-FSSHTTP 2.2.3.7
                new XElement(
                    soap + "ResponseVersion",
                    new XAttribute("Version", "2"),
                    new XAttribute("MinorVersion", minorVersion)),
                // MS-FSSHTTP 2.2.3.6
                new XElement(
                    soap + "ResponseCollection",
                    new XAttribute("WebUrl", baseUrl + site.Path),
                    new XAttribute("WebUrlIsEncoded", "false"),
                    responses)));
        var soapXml = body.ToString(SaveOptions.DisableFormatting);
        var response = context.Response;
        // MS-FSSHTTP 2.1: the server always answers in MTOM
        var boundary = $"urn:uuid:{Guid.NewGuid()}";
        response.ContentType =
            $"multipart/related; type=\"application/xop+xml\"; boundary=\"{boundary}\"; start=\"<http://tempuri.org/0>\"; start-info=\"text/xml\"";
        await response.WriteAsync(
            $"--{boundary}\r\n" +
            "Content-ID: <http://tempuri.org/0>\r\n" +
            "Content-Transfer-Encoding: 8bit\r\n" +
            "Content-Type: application/xop+xml;charset=utf-8;type=\"text/xml\"\r\n" +
            "\r\n" +
            soapXml +
            $"\r\n--{boundary}--\r\n");
    }

    /// <summary>
    /// The Response of a Request ([MS-FSSHTTP] 2.2.3.5), with a SubResponse per SubRequest (2.2.3.10).
    /// </summary>
    static XElement Response(XElement request, string url, Site site, WopiFile? file)
    {
        var results = new Dictionary<string, string>();
        var subResponses = new List<XElement>();
        foreach (var subRequest in request.Elements(soap + "SubRequest"))
        {
            var token = (string?) subRequest.Attribute("SubRequestToken") ?? "";
            var (errorCode, data) = SubResponse(subRequest, results, site, file);
            results[token] = errorCode;
            subResponses.Add(
                new(
                    soap + "SubResponse",
                    new XAttribute("SubRequestToken", token),
                    new XAttribute("ErrorCode", errorCode),
                    new XAttribute("HResult", errorCode == "Success" ? 0 : failHResult),
                    data));
        }

        return new(
            soap + "Response",
            new XAttribute("Url", url),
            new XAttribute("UrlIsEncoded", "true"),
            new XAttribute("RequestToken", (string?) request.Attribute("RequestToken") ?? "1"),
            new XAttribute("HealthScore", "0"),
            subResponses);
    }

    static async Task<WopiFile?> FileOfUrl(string url, string baseUrl, Site site, FileSystemStorage storage)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var path = uri.AbsolutePath;
        var prefix = new Uri(baseUrl + site.Path).AbsolutePath.TrimEnd('/');
        if (prefix.Length > 0 &&
            path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            path = path[prefix.Length..];
        }

        return await FindFile(storage, path);
    }

    /// <summary>
    /// The error code and SubResponseData of a SubRequest ([MS-FSSHTTP] 3.1.4).
    /// </summary>
    static (string ErrorCode, XElement Data) SubResponse(
        XElement subRequest,
        Dictionary<string, string> results,
        Site site,
        WopiFile? file)
    {
        var type = (string?) subRequest.Attribute("Type") ?? "";
        var requestData = subRequest.Element(soap + "SubRequestData");
        var data = new XElement(soap + "SubResponseData");

        // MS-FSSHTTP 2.2.5.3: OnNotSupported runs only when the subrequest it depends on was not supported
        var dependsOn = (string?) subRequest.Attribute("DependsOn");
        var dependencyType = (string?) subRequest.Attribute("DependencyType");
        if (dependsOn != null &&
            dependencyType == "OnNotSupported" &&
            results.TryGetValue(dependsOn, out var dependency) &&
            dependency != "RequestNotSupported")
        {
            // MS-FSSHTTP 2.2.5.2, as in the example of 4.1.2
            return ("DependentOnlyOnNotSupportedRequestGetSupported", data);
        }

        switch (type)
        {
            case "WhoAmI":
                // MS-FSSHTTP 3.1.4.6, 2.3.3.6
                var user = site.UserId ?? "anonymous";
                data.Add(
                    new XAttribute("UserName", site.UserName ?? user),
                    new XAttribute("UserLogin", $"spike\\{user}"),
                    new XAttribute("UserEmailAddress", $"{user}@example.com"),
                    new XAttribute("UserSIPAddress", $"{user}@example.com"));
                return ("Success", data);
            case "ServerTime":
                // MS-FSSHTTP 3.1.4.7, 2.3.1.18: in ticks
                data.Add(new XAttribute("ServerTime", DateTime.UtcNow.Ticks));
                return ("Success", data);
            case "GetDocMetaInfo":
                // MS-FSSHTTP 3.1.4.9, 2.3.1.27
                if (file == null)
                {
                    return ("FileNotExistsOrCannotBeCreated", data);
                }

                data.Add(
                    new XElement(
                        soap + "DocProps",
                        Property("vti_title", Path.GetFileNameWithoutExtension(file.Name)),
                        Property("vti_filesize", file.Size.ToString()),
                        Property("vti_timelastmodified", file.LastModified.UtcDateTime.ToString("O")),
                        Property("vti_etag", $"\"{file.Version}\"")),
                    new XElement(
                        soap + "FolderProps",
                        Property("vti_title", "spike")));
                return ("Success", data);
            case "Coauth":
                // MS-FSSHTTP 3.1.4.3: stubbed, always alone
                var coauthType = (string?) requestData?.Attribute("CoauthRequestType");
                if (coauthType is "JoinCoauthoring" or "RefreshCoauthoring")
                {
                    data.Add(
                        new XAttribute("LockType", "SchemaLock"),
                        new XAttribute("CoauthStatus", "Alone"),
                        new XAttribute("TransitionID", file?.Id ?? Guid.Empty.ToString()));
                }
                else if (coauthType == "GetCoauthoringStatus")
                {
                    data.Add(new XAttribute("CoauthStatus", "Alone"));
                }

                return ("Success", data);
            case "SchemaLock":
                // MS-FSSHTTP 3.1.4.4: stubbed
                var schemaLockType = (string?) requestData?.Attribute("SchemaLockRequestType");
                if (schemaLockType is "GetLock" or "RefreshLock")
                {
                    data.Add(new XAttribute("LockType", "SchemaLock"));
                }

                return ("Success", data);
            case "ExclusiveLock":
            case "EditorsTable":
                // MS-FSSHTTP 3.1.4.5, 3.1.4.8: stubbed
                return ("Success", data);
            default:
                // Cell (MS-FSSHTTP 3.1.4.2) needs FSSHTTPB, which the spike does not have
                return ("RequestNotSupported", data);
        }
    }

    /// <summary>
    /// A property of GetDocMetaInfo ([MS-FSSHTTP] 2.3.1.29).
    /// </summary>
    static XElement Property(string key, string value) =>
        new(soap + "Property", new XAttribute("Key", key), new XAttribute("Value", value));

    #endregion
}
