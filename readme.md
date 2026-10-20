#  <img src="/src/icon.png" height="30px"> OfficeHost

[![Build](https://github.com/Papyrine/OfficeHost/actions/workflows/build.yml/badge.svg)](https://github.com/Papyrine/OfficeHost/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/OfficeHost.svg?label=OfficeHost)](https://www.nuget.org/packages/OfficeHost/)

The host side of the Office document protocols as a library for ASP.NET Core (.NET 10), formerly the WopiServer
package. Add it to a web application, and clients of the Web Application Open Platform Interface
([WOPI](https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/)) (Collabora Online,
ONLYOFFICE, Microsoft 365 for the web) open, edit and save the files of any storage the application has.


## How it works

1. A user opens a document in the application. The application asks `OfficeLauncher` for the URL of the editor and a WOPI
   access token, scoped to that file and user.
2. The page of the application posts the token to the editor URL, usually in an iframe.
3. The editor calls the WOPI endpoints the library maps (`/wopi/files/{fileId}`) to read the file metadata, lock the
   file, and get and save its content. The library checks the token and the locks, and goes to the storage of the
   application through `IOfficeStorage`.


## Features

* The WOPI file operations: CheckFileInfo, GetFile, PutFile, Lock, GetLock, RefreshLock, Unlock, UnlockAndRelock,
  PutRelativeFile, RenameFile, DeleteFile, PutUserInfo. Conformance to [MS-WOPI] is tested section by section, and
  Microsoft's WOPI validator passes.
* Access tokens that grant a user's access to one file: signed tokens out of the box, or any token scheme through
  `IOfficeTokens`.
* The discovery of the WOPI clients, to launch their actions and validate the proof keys of their requests
  (`X-WOPI-Proof`).
* Pluggable storage and locks: `FileSystemStorage` and `MemoryLockStore` are included.
* Desktop Office, in preview: Word opens, edits and saves files through the cell storage service of SharePoint
  ([MS-FSSHTTP]), with the shared lock and coauthoring. See [Desktop Office](#desktop-office).


## Usage

The code below is from [the sample](#sample), which imports the `OfficeHost` namespace. Register a storage and the WOPI
services, map the WOPI endpoints, and add the pages that open a file:

<!-- snippet: Registration -->
<a id='snippet-Registration'></a>
```cs
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
app.MapFssHttp();
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
```
<sup><a href='/src/OfficeHost.Sample/Program.cs#L11-L71' title='Snippet source file'>snippet source</a> | <a href='#snippet-Registration' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

A page decides what the user may do with the file, gets the URL of the editor and an access token from `OfficeLauncher`,
and returns a page that posts the token to the editor in an iframe
([`SamplePages.Editor`](src/OfficeHost.Sample/SamplePages.cs)):

<!-- snippet: Launch -->
<a id='snippet-Launch'></a>
```cs
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
```
<sup><a href='/src/OfficeHost.Sample/Program.cs#L82-L124' title='Snippet source file'>snippet source</a> | <a href='#snippet-Launch' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The WOPI clients call back the application: they must reach it at the URL of the request that launched the editor, or
at `BaseUrl` when it differs, as when the editor runs in a container or behind a reverse proxy.


## Sample

[src/OfficeHost.Sample](src/OfficeHost.Sample) is this setup as a running application: it lists the files of a folder
and opens them in Collabora Online, with a host page that loads the editor in an iframe. To try it, start Collabora
Online (CODE) in Docker, allowing it to call back the sample. Home mode turns off its welcome dialog, and caps it at 20
connections and 10 documents, which is plenty for trying the sample:

```pwsh
docker run --rm --publish 9980:9980 --add-host host.docker.internal:host-gateway --env "domain=host\.docker\.internal.*" --env "extra_params=--o:ssl.enable=false --o:ssl.termination=false --o:home_mode.enable=true" collabora/code
```

Then run the sample, and open http://localhost:5000:

```pwsh
dotnet run --project src/OfficeHost.Sample
```

It starts with a `Welcome.docx` in `src/OfficeHost.Sample/files`; add files there to open them too. Edit opens a file
with write access, View read only. The URLs and the signing key are in its `appsettings.json`: `OfficeHost:BaseUrl` is how
Collabora reaches the sample (`host.docker.internal` from the container), and `Sample:AppUrl` how the browser does.


## Options

| Option | Default | |
|---|---|---|
| `SigningKey` | | Signs the access tokens, at least 32 characters. Not needed with another `IOfficeTokens`. |
| `TokenLifetime` | 10 hours | How long an access token is valid. |
| `LockDuration` | 30 minutes | How long a lock lasts unless refreshed, as [MS-WOPI] sets it. |
| `BaseUrl` | the request's | The URL the WOPI clients reach the application at: WOPISrc is `{BaseUrl}/wopi/files/{fileId}`. |
| `DiscoveryUrls` | | The trusted WOPI clients: their actions are launched, and their proof keys checked. |
| `RequireProof` | false | Refuses requests without `X-WOPI-Proof`. Collabora Online does not sign its requests. |
| `MaxFileSize` | no limit | The largest file PutFile and PutRelativeFile accept, replacing the server's request size limit. |
| `ServerVersion` | the library's | The `X-WOPI-ServerVersion` of the responses. |
| `HostEditUrl`, `HostViewUrl` | | The pages of the application that open a file, returned by CheckFileInfo and PutRelativeFile. |
| `OnCheckFileInfo` | | Adds or changes CheckFileInfo properties: `CloseUrl`, breadcrumbs, branding, CSPP+ properties. |
| `DesktopPath` | `/fsshttp` | The path of the sites of desktop Office, `MapFssHttp`. |
| `DesktopBaseUrl` | the request's | The URL desktop Office reaches the application at, which runs where the browser does. |


## Access

`OfficeAccess` is what a token grants: a user's access to a file, which the application decides when it launches the
editor.

* `CanWrite`: lock, save and delete the file. Without it the file opens read only.
* `CanWriteRelative`: Save As, creating files next to this one.
* `CanRename`: rename the file.
* `ViewOnly`: no download, print or copy.
* `IsAnonymous`, `UserName`: how the editor shows the user.

The built-in tokens carry the access, signed with `SigningKey`, so they need no storage, but cannot be revoked before
they expire. To keep tokens in a database, or tie them to the application's sessions, register an `IOfficeTokens`.


## Storage

The storage is the application's: implement `IOfficeStorage` over a database, blob storage or a document management
system. File ids are chosen by the storage, and must not change when a file is saved or renamed, as WOPI clients keep
using them. A missing file throws `FileNotFoundException`. The storage, the stores below and `IOfficeTokens` can be
registered with any lifetime, as scoped for a storage over a database context.

```csharp
public interface IOfficeStorage
{
    Task<OfficeFile?> GetFile(string fileId, Cancel cancel = default);
    Task<OfficeFile?> GetSibling(string fileId, string name, Cancel cancel = default);
    Task<Stream> OpenRead(string fileId, Cancel cancel = default);
    Task<OfficeFile> Write(string fileId, Stream content, Cancel cancel = default);
    Task<OfficeFile?> CreateSibling(string fileId, string name, Stream content, bool overwrite, Cancel cancel = default);
    Task<OfficeFile?> Rename(string fileId, string name, Cancel cancel = default);
    Task Delete(string fileId, Cancel cancel = default);
}
```

`FileSystemStorage` keeps the files in a folder, for development, tests and single machines. It gives each file an id
on first use (`GetFileId("reports/q3.docx")`), kept in a hidden `.wopi` folder so that ids survive renames, and writes
through a temporary file so that a save never leaves a partial file.

Locks are kept by an `IOfficeLockStore`. The default `MemoryLockStore` suits a single instance of the application;
several instances need a shared store, such as one over Redis or a database, with an atomic compare and exchange.

PutUserInfo, where the editors keep per-user data such as the tips a user has seen, is supported when an
`IOfficeUserInfoStore` is registered.

The state of desktop coauthoring (the data elements Office syncs, the coauthoring session) is kept by an
`IFssHttpStore`: the default `MemoryFssHttpStore` for a single instance, `FileSystemFssHttpStore` for a single machine,
or a shared store with an atomic compare and exchange. The file stays in the storage: a lost state only makes Office
download the file again. A WOPI DeleteFile deletes the state of the file too; an application that deletes files itself
calls `IFssHttpStore.Delete` after them, or the state stays, unused.


## Desktop Office

Preview. Desktop Word opens files as it opens those of a SharePoint site: through the cell storage service of
[MS-FSSHTTP], which syncs a file in chunks ([MS-FSSHTTPB], [MS-FSSHTTPD]), shares the lock between the users of a
coauthoring session, and keeps an editors table.

* `MapFssHttp()` maps the sites: `{DesktopPath}/{token}/{name}` is the URL of a file, the token of `IOfficeTokens` in its
  path standing for the file and the user, as Office blocks Basic and forms-based sign-in by default.
* `OfficeLauncher.LaunchDesktop(access, request)` gives the link that opens the file: `ms-word:ofe|u|{url}` (or
  `ms-excel:`, `ms-powerpoint:`; `ofv` without write access).
* The lock of desktop editing is a lock of the `IOfficeLockStore`, so WOPI clients and desktop Office do not edit a file
  at the same time. A file changed outside desktop Office is synced as a new revision, keeping the unchanged chunks.

The host must use HTTPS with a certificate Windows trusts. A host in the Internet zone opens in Protected View first, as
SharePoint Online does. Tested with Word for Microsoft 365 16.0.20326 on Windows. Notes on what Word does that the
specs do not say are in [docs/fsshttp-notes.md](docs/fsshttp-notes.md), and the check by hand with Word in
[docs/fsshttp-checklist.md](docs/fsshttp-checklist.md).


## Proof keys

WOPI clients that publish a `proof-key` in their discovery sign their requests ([MS-WOPI] 2.2.1). The requests that
carry `X-WOPI-Proof` are validated against the keys of `DiscoveryUrls`: the current and the old key, for clients that
rotate theirs. The keys are cached for an hour, and refetched on a mismatch when over a minute old, so forged requests
do not turn into a discovery request each. When a client is down, its last keys are kept, and its discovery is attempted
again once a minute.

Requests without proof are accepted unless `RequireProof` is set. The signed URL is `BaseUrl` (or the request's) and the
path and query of the request, so `BaseUrl` must be the URL the clients call.


## Development

```pwsh
dotnet build src
dotnet run --project src/OfficeHost.Tests --no-build
```

The test project is a TUnit executable: run it, rather than `dotnet test`. It holds:

* `SpecTests`: conformance to the server side of [MS-WOPI], each test citing its section.
* `EdgeCaseTests`: protocol edge cases ported from [WopiHost](https://github.com/petrsvihlik/WopiHost).
* `ProofKeyTests`: the proof key validation, including key rotation.
* `FailureTests`: storage failures, through a storage that fails on demand.
* `StorageTests`: `FileSystemStorage` and `MemoryLockStore`.
* `ApiTests`: the launcher, the tokens, the options and the registration.
* `KestrelTests`: the request size limits, on Kestrel.
* `SampleTests`: the sample application, with a stubbed WOPI client.

`src/OfficeHost.E2ETests` hosts the endpoints on a real port and opens documents in real editors, Collabora and
ONLYOFFICE, in Docker, driven with Playwright: Collabora opens, edits and saves a document, and ONLYOFFICE opens one
with every request validated against its proof keys. It needs Docker and the Playwright browser, and pulls the editor
images on the first run (about 1.5 and 4.3 GB):

```pwsh
dotnet build src/OfficeHost.E2ETests
pwsh src/OfficeHost.E2ETests/bin/Debug/net10.0/playwright.ps1 install chromium
dotnet run --project src/OfficeHost.E2ETests --no-build -- --treenode-filter '/*/*/EditorTests/*'
```

Microsoft's [WOPI validator](https://github.com/microsoft/wopi-validator-core) runs with
`pwsh src/wopi-validator.ps1` (`-Configuration Debug` for a Debug build): it builds the validator, then runs
`ValidatorTests` of the same project, each test group of the validator on a fresh storage. The ProofKeys group is left
out: the validator's command line sends no proofs, so `ProofKeyTests` covers them instead.

Code coverage of the library, with a per-file table:

```pwsh
dotnet run --project src/OfficeHost.Tests -- --coverage --coverage-output-format cobertura --coverage-output coverage.cobertura.xml --coverage-settings $PWD/src/coverage.config --results-directory $PWD/TestResults
pwsh src/coverage-summary.ps1 TestResults/coverage.cobertura.xml
```

`-Filter FssHttp/Binary/` limits the table and the total to the files of a folder, here the FSSHTTPB codec.

CI builds and runs the tests on Linux and Windows, with the coverage in the job summary, and runs the WOPI validator on
Linux. The editor tests run nightly and on demand (`e2e.yml`), not on every push.

A Release build packs the library into `nugets/`. Pushing a tag releases it: the tag is the version (`1.2.3` or
`v1.2.3`), and the `publish` job of `build.yml` pushes the package to nuget.org through trusted publishing, once the
tests and the validator pass.


## Credits and license

The WOPI logic started as a port of the Python [CS3 WOPI server](https://github.com/cs3org/wopiserver), by Giuseppe Lo
Presti (CERN) and its contributors. Like it, this library is under the Apache 2.0 license, see
[license.txt](license.txt).

[MS-WOPI]: https://learn.microsoft.com/en-us/openspecs/office_protocols/ms-wopi/
[MS-FSSHTTP]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/
[MS-FSSHTTPD]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpd/


## Icon

https://thenounproject.com/icon/fountain-pen-8154546/
