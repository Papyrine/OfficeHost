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
services.AddSingleton<IFsshttpStore>(new FileSystemFsshttpStore(states));
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

var app = builder.Build();
// the WOPI endpoints the WOPI client calls: /wopi/files/{fileId}
app.MapOfficeHost();
// the pages that open a file in the WOPI client
app.MapGet(
    "/edit/{fileId}",
    (string fileId, HttpRequest request, OfficeLauncher launcher) =>
        Open(fileId, "edit", request, launcher));
app.MapGet(
    "/view/{fileId}",
    (string fileId, HttpRequest request, OfficeLauncher launcher) =>
        Open(fileId, "view", request, launcher));
// the sites desktop Office edits and coauthors files through: /fsshttp/{token}/{name}
app.MapFsshttp();
// the page that opens a file in desktop Office, through an ms-word: (or ms-excel:, ms-powerpoint:) link; the user comes
// from ?user= so that two people can coauthor the demo
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
            CanWrite = true
        };
        var link = await launcher.LaunchDesktop(access, request);
        return Results.Content(SamplePages.Desktop(link), "text/html");
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
static async Task<IResult> Open(
    string fileId,
    string action,
    HttpRequest request,
    OfficeLauncher launcher)
{
    // the access of the user to the file: an application decides it from its own users and permissions
    var canWrite = action == "edit";
    var access = new OfficeAccess(fileId, "demo")
    {
        UserName = "Demo User",
        CanWrite = canWrite,
        CanWriteRelative = canWrite,
        CanRename = canWrite
    };
    try
    {
        var launch = await launcher.Launch(access, action, request);
        return Results.Content(SamplePages.Editor(launch), "text/html");
    }
    catch (FileNotFoundException)
    {
        return Results.NotFound();
    }
    catch (OfficeClientUnreachableException exception)
    {
        // the WOPI client is down, or not started
        return Results.Content(
            SamplePages.Unreachable(exception.Message),
            "text/html",
            statusCode: StatusCodes.Status503ServiceUnavailable);
    }
    catch (InvalidOperationException exception)
    {
        // no WOPI client offers the action for this type of file
        return Results.Content(
            SamplePages.Message(exception.Message),
            "text/html",
            statusCode: StatusCodes.Status400BadRequest);
    }
}
// end-snippet
