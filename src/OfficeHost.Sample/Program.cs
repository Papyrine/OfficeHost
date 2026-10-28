// A web application that opens the files of a folder in a WOPI client, such as Collabora Online.
var builder = WebApplication.CreateBuilder(args);
var sample = builder.Configuration.GetSection("Sample");
var appUrl = sample["AppUrl"]!.TrimEnd('/');
var folder = Path.Combine(builder.Environment.ContentRootPath, sample["Files"]!);
Directory.CreateDirectory(folder);
SampleFiles.AddWelcome(folder);
// a folder to log every request and response to, as the desktop launch profile sets, to see what desktop Office sends
var requestLog = sample["RequestLog"];

// begin-snippet: Registration
// the storage of the files, here a folder: implement IOfficeStorage for any other
var storage = new FileSystemStorage(folder);
var services = builder.Services;
services.AddSingleton(storage);
services.AddSingleton<IOfficeStorage>(storage);
// the state of desktop coauthoring, kept in a folder so that a restart does not make desktop Office download it all
var states = Path.Combine(builder.Environment.ContentRootPath, "fsshttp-state");
services.AddSingleton<IFssHttpStore>(new FileSystemFssHttpStore(states));
services.AddOfficeHost(
    options =>
    {
        // SigningKey, BaseUrl and DiscoveryUrls come from appsettings.json
        builder.Configuration.GetSection("OfficeHost").Bind(options);
        // the pages below, which the WOPI client links to, as after a Save As
        options.HostEditUrl = (file, _) => $"{appUrl}/edit/{Uri.EscapeDataString(file.Id)}";
        options.HostViewUrl = (file, _) => $"{appUrl}/view/{Uri.EscapeDataString(file.Id)}";
        options.OnCheckFileInfo = context =>
        {
            // back to the file list when the user closes the editor
            context.Properties["CloseUrl"] = appUrl;
            context.Properties["BreadcrumbBrandName"] = "OfficeHost sample";
            context.Properties["BreadcrumbBrandUrl"] = appUrl;
            return Task.CompletedTask;
        };
    });

// begin-snippet: Shadow
// a shadow of SharePoint (docs/shadow.md), when the OfficeHostShadow section has an ApiKey:
// the module in SharePoint posts desktop Office's traffic here, and the shadow replays it
// to this application's OfficeHost and logs how the answers compare. A deployment of its own,
// with storage of its own: SharePoint's documents are written to it
var shadow = !string.IsNullOrEmpty(builder.Configuration["OfficeHostShadow:ApiKey"]);
if (shadow)
{
    // the documents of SharePoint as files of the folder, under their host and path
    services.AddSingleton<IShadowDocuments>(new FileSystemShadowDocuments(storage));
    // the documents Office opened from its cache, from SharePoint, as the application's account
    services.AddHttpClient<IShadowContent, SharePointShadowContent>()
        .ConfigurePrimaryHttpMessageHandler(
            () => new HttpClientHandler
            {
                UseDefaultCredentials = true
            });
    services.AddOfficeHostShadow();
}
// end-snippet

var app = builder.Build();
if (shadow)
{
    // where the module posts: /officehost-shadow/exchanges
    app.MapOfficeHostShadow();
}

// the WOPI endpoints the WOPI client calls: /wopi/files/{fileId}
app.MapOfficeHost();
// the pages that open a file in the WOPI client: /edit/{fileId} and /view/{fileId}
app.MapOfficePages(Access, OnError);
// the sites desktop Office edits and coauthors files through: /fsshttp/{token}/{name}
app.MapFssHttp();
// the page that opens a file in desktop Office, through an ms-word: (or ms-excel:, ms-powerpoint:) link. The user comes
// from ?user= so that two people can coauthor the demo, and may edit any file: an application takes the user from its
// sign-in, and decides the access from its permissions, as Access below does for the WOPI client
app.MapGet(
    "/desktop/{fileId}",
    async (string fileId, string? user, HttpRequest request, OfficeLauncher launcher) =>
    {
        if (string.IsNullOrEmpty(user))
        {
            user = "demo";
        }

        var access = new OfficeAccess(fileId, user)
        {
            UserName = $"{char.ToUpperInvariant(user[0])}{user[1..]} User",
            CanWrite = true,
            // Save As to a new name in the folder
            CanWriteRelative = true,
            // the rename of Office's title bar
            CanRename = true
        };
        var link = await launcher.LaunchDesktop(access, request);
        return Results.Content(SamplePages.Desktop(link), "text/html");
    });
// end-snippet

// begin-snippet: NewDocument
// a new document: an empty file, which the WOPI client (its editnew action) or desktop Office fills on the first save
app.MapPost(
    "/new",
    async (HttpRequest request) =>
    {
        var form = await request.ReadFormAsync();
        var name = SampleFiles.NewName(folder, form["type"].ToString());
        await File.WriteAllBytesAsync(Path.Combine(folder, name), []);
        var fileId = Uri.EscapeDataString(storage.GetFileId(name));
        var page = $"/editnew/{fileId}";
        if (form["open"] == "desktop")
        {
            page = $"/desktop/{fileId}";
        }

        return Results.Redirect(page);
    });
// end-snippet
if (!string.IsNullOrEmpty(requestLog))
{
    app.UseRequestLog(Path.Combine(builder.Environment.ContentRootPath, requestLog));
}

app.MapGet("/", () => Results.Content(SamplePages.Files(storage), "text/html"));
// the server listens on 0.0.0.0 for the WOPI client in Docker, which is not an address to browse
app.Lifetime.ApplicationStarted.Register(() => app.Logger.LogInformation("Open {AppUrl} in a browser", appUrl));
app.Run();

// begin-snippet: Launch
// the access of the user to the file: an application decides it from its own users and permissions, or returns null to
// refuse it
static ValueTask<OfficeAccess?> Access(HttpContext context, string fileId, string action)
{
    var canWrite = action is "edit" or "editnew";
    var access = new OfficeAccess(fileId, "demo")
    {
        UserName = "Demo User",
        CanWrite = canWrite,
        CanWriteRelative = canWrite,
        CanRename = canWrite
    };
    return ValueTask.FromResult<OfficeAccess?>(access);
}

// a page of the sample when the WOPI client is down, rather than the default 503 text
static IResult? OnError(HttpContext context, Exception exception)
{
    if (exception is OfficeClientUnreachableException)
    {
        return Results.Content(
            SamplePages.Unreachable(exception.Message),
            "text/html",
            statusCode: StatusCodes.Status503ServiceUnavailable);
    }

    return null;
}
// end-snippet
