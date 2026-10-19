// A web application that opens the files of a folder in a WOPI client, such as Collabora Online.
var builder = WebApplication.CreateBuilder(args);
var sample = builder.Configuration.GetSection("Sample");
var appUrl = sample["AppUrl"]!.TrimEnd('/');
var folder = Path.Combine(builder.Environment.ContentRootPath, sample["Files"]!);
Directory.CreateDirectory(folder);
SampleFiles.AddWelcome(folder);

// begin-snippet: Registration
// the storage of the files, here a folder: implement IWopiStorage for any other
var storage = new FileSystemStorage(folder);
builder.Services.AddSingleton(storage);
builder.Services.AddSingleton<IWopiStorage>(storage);
builder.Services.AddWopi(
    options =>
    {
        // SigningKey, BaseUrl and DiscoveryUrls come from appsettings.json
        builder.Configuration.GetSection("Wopi").Bind(options);
        // the pages below, which the WOPI client links to, as after a Save As
        options.HostEditUrl = (file, _) => $"{appUrl}/edit/{Uri.EscapeDataString(file.Id)}";
        options.HostViewUrl = (file, _) => $"{appUrl}/view/{Uri.EscapeDataString(file.Id)}";
        options.OnCheckFileInfo = context =>
        {
            // back to the file list when the user closes the editor
            context.Properties["CloseUrl"] = appUrl;
            context.Properties["BreadcrumbBrandName"] = "WopiServer sample";
            context.Properties["BreadcrumbBrandUrl"] = appUrl;
            return Task.CompletedTask;
        };
    });

var app = builder.Build();
// the WOPI endpoints the WOPI client calls: /wopi/files/{fileId}
app.MapWopi();
// the pages that open a file in the WOPI client
app.MapGet(
    "/edit/{fileId}",
    (string fileId, HttpRequest request, WopiLauncher launcher) => Open(fileId, "edit", request, launcher));
app.MapGet(
    "/view/{fileId}",
    (string fileId, HttpRequest request, WopiLauncher launcher) => Open(fileId, "view", request, launcher));
// end-snippet
app.MapGet("/", () => Results.Content(SamplePages.Files(storage), "text/html"));
// the server listens on 0.0.0.0 for the WOPI client in Docker, which is not an address to browse
app.Lifetime.ApplicationStarted.Register(() => app.Logger.LogInformation("Open {AppUrl} in a browser", appUrl));
app.Run();

// begin-snippet: Launch
static async Task<IResult> Open(string fileId, string action, HttpRequest request, WopiLauncher launcher)
{
    // the access of the user to the file: an application decides it from its own users and permissions
    var canWrite = action == "edit";
    var access = new WopiAccess(fileId, "demo")
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
    catch (WopiClientUnreachableException exception)
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
