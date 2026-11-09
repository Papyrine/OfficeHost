// A shadow of SharePoint on SQL Server (docs/shadow.md): the module in SharePoint posts desktop Office's traffic here,
// and the shadow replays it to OfficeHost, whose files, locks and coauthoring state are in a database of its own, and
// logs how its answers compare with SharePoint's. Office goes on working with SharePoint: no one opens files here.
var builder = WebApplication.CreateBuilder(args);
var configuration = builder.Configuration;
var services = builder.Services;
// a Windows service when started as one, a console application otherwise
services.AddWindowsService(options => options.ServiceName = "OfficeHost Shadow");

var connectionString = configuration.GetConnectionString("OfficeHost") ??
                       throw new InvalidOperationException("No connection string OfficeHost");
var schema = configuration["Shadow:Schema"] ?? "dbo";
if (configuration.GetValue("Shadow:InstallDatabase", true))
{
    await ShadowDatabase.Install(connectionString, schema);
}

// begin-snippet: ShadowSqlServer
// the files, locks and coauthoring state of OfficeHost in SQL Server, as the application would keep them
services.AddOfficeHostSqlServer(connectionString, schema);
services.AddOfficeHost(options => configuration.GetSection("OfficeHost").Bind(options));
// the documents of SharePoint as files of the database, in the folder of their host and path
services.AddSingleton<IShadowDocuments>(_ => new SqlServerShadowDocuments(_.GetRequiredService<SqlServerStorage>()));
// the documents Office opened from its cache, or downloaded in parts, from SharePoint, as the service's account
services.AddHttpClient<IShadowContent, SharePointShadowContent>()
    .ConfigurePrimaryHttpMessageHandler(
        () => new HttpClientHandler
        {
            UseDefaultCredentials = true
        });
services.AddOfficeHostShadow();
// end-snippet

var app = builder.Build();
// to check that the servers of SharePoint reach the shadow
app.MapGet("/", () => "OfficeHost shadow: the module in SharePoint posts to /officehost-shadow/exchanges");
app.MapOfficeHostShadow();
app.Run();
