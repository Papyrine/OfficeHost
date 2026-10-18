# WopiServer

[![Build](https://github.com/Papyrine/WopiServer/actions/workflows/build.yml/badge.svg)](https://github.com/Papyrine/WopiServer/actions/workflows/build.yml)

A WOPI server for ASP.NET Core (.NET 10). It implements the host side of the Web Application Open Platform Interface
([WOPI](https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/)), so a file storage can
offer online editing with Microsoft 365 for the web, Collabora Online or ONLYOFFICE. A bridge also integrates editors
without WOPI support: CodiMD and Etherpad.

It is a port of the Python [CS3 WOPI server](https://github.com/cs3org/wopiserver), and uses the same configuration
file. See [the porting notes](docs/porting.md) for what was ported and how the two differ.


## How it works

1. A user opens a document in your file storage (EFSS). The storage calls `/wopi/iop/openinapp` with a shared secret,
   naming the file, the user and the editor.
2. The server checks the file, and returns the URL of the editor and a WOPI access token: a JWT scoped to that file,
   user and view mode.
3. The storage opens the editor with the token. The editor calls the WOPI endpoints of this server
   (`/wopi/files/<id>`) to read the file metadata, lock the file, and get and save its contents.


## Features

* The WOPI file operations: CheckFileInfo, GetFile, PutFile, Lock, GetLock, RefreshLock, Unlock, UnlockAndRelock,
  PutRelativeFile, RenameFile, DeleteFile, PutUserInfo. Conformance to [MS-WOPI] is tested section by section.
* Locks that interoperate with desktop apps: Microsoft Office and LibreOffice lock files block editing online, and a
  LibreOffice-compatible lock file is created while a file is edited online.
* Detection of files changed behind the server's back, with conflict copies instead of lost changes, and a recovery
  space for saves that fail.
* Validation of the proof keys of the WOPI clients (`X-WOPI-Proof`).
* The bridge for CodiMD and Etherpad, which saves the documents back to the storage as they change.
* Pluggable storage: a local folder storage is included, any other plugs in through `IStorage`.
* JSON lines logs and Prometheus metrics.


## Running

```pwsh
dotnet run --project src/WopiServer -- --config /path/to/wopiserver.conf
```

The configuration is loaded in this order, later files overriding earlier ones:

1. `wopiserver.defaults.conf`, shipped next to the executable. It documents every option.
2. `/etc/wopi/wopiserver.defaults.conf`, when present, or the file given with `--defaults`.
3. `/etc/wopi/wopiserver.conf`, or the file given with `--config`. This file is watched: options read per request (log
   level, URLs advertised to the editors, lock and token lifetimes...) apply without a restart, while the storage, the
   port, `wopiurl`, HTTPS and the secret files are read at startup.

A minimal configuration over a local folder:

```ini
[general]
storagetype = local
wopiurl = http://localhost:8880
loghandler = stream

[security]
wopisecretfile = /etc/wopi/wopisecret
iopsecretfile = /etc/wopi/iopsecret

[io]
recoverypath = /var/spool/wopirecovery

[local]
storagehomepath = /var/wopi_local_storage
```

The two secret files hold random strings: `wopisecret` signs the access tokens, `iopsecret` is the shared secret of
the calls from the storage. Then open a file for an editor:

```pwsh
curl -H "Authorization: Bearer <iopsecret>" -H "TokenHeader: x" "http://localhost:8880/wopi/iop/openinapp?fileid=/test.docx&viewmode=VIEW_MODE_READ_WRITE&appname=MSOffice&appurl=https://your-office/edit&username=Alice"
```

The response gives the URL to open the editor with, and the token to post to it as `access_token`:

```json
{"app-url":"https://your-office/edit?WOPISrc=http%3A%2F%2Flocalhost%3A8880%2Fwopi%2Ffiles%2Flocal%21...","form-parameters":{"access_token":"eyJ..."}}
```


## Configuration

The main options, by section. `wopiserver.defaults.conf` describes them all, with their defaults.

| Option | |
|---|---|
| **[general]** | |
| `storagetype` | The storage: `local`. When embedding with a custom `IStorage`, it only names the storage on the information page. |
| `port`, `wopiurl` | The port to listen on, and the public URL of the server, as WOPI clients reach it. |
| `loglevel`, `loghandler`, `logdest` | `Debug` to `Critical`; `file` (to `logdest`) or `stream` (`stdout` or `stderr`). |
| `hostediturl`, `hostviewurl` | The pages of the storage that embed the editor, with `<path>`, `<endpoint>`, `<fileid>` and `<app>` placeholders. |
| `downloadurl`, `filesharingurl`, `webdavurl` | Optional URLs advertised to the editors. |
| `brandingname`, `brandingurl` | The name and link of the storage shown by the editors. |
| `tokenvalidity` | The lifetime of the access tokens, in seconds (1 day). |
| `wopilockexpiration` | The lifetime of the locks, in seconds (30 minutes, as WOPI requires). |
| `detectexternallocks`, `detectexternalmodifications` | The interoperability with desktop apps, on by default. |
| `homepath` | Where conflict copies and "save as" go when the folder of the file is not writable. |
| `enablerename` | Allows RenameFile, off by default (see [the porting notes](docs/porting.md)). |
| `noversiongraceperiod` | Skips creating a new version for saves within this many seconds. |
| `wopiproxy`, `wopiproxysecretfile`, `proxiedappname` | Routes an app through ownCloud's WOPI proxy. |
| **[security]** | |
| `wopisecretfile`, `iopsecretfile` | The secret files. |
| `usehttps`, `wopicert`, `wopikey` | Serves HTTPS with a PEM certificate and key. |
| `proofkeydiscoveryurls`, `proofkeyrequired` | See [Proof keys](#proof-keys). |
| **[apps]** | |
| `codeofficetypes`, `nonofficetypes` | The file types that get a LibreOffice lock file, and those that get no desktop lock checks. |
| `businessflow`, `earlyfeatures`, `compliancedomain` | Microsoft 365 settings. |
| **[io]** | |
| `chunksize`, `recoverypath` | The read buffer size, and the folder of the recovery space. |
| **[local]** | |
| `storagehomepath` | The folder of the local storage. |
| **[bridge]** | |
| `sslverify`, `saveinterval`, `unlockinterval`, `disablezip` | The bridge's certificate check, save and unlock intervals, and CodiMD bundles. |


## Endpoints

| Endpoint | |
|---|---|
| `GET /wopi/iop/openinapp` | Opens a file for an editor: returns the editor URL and an access token. Needs `Authorization: Bearer <iopsecret>`. |
| `GET /wopi/iop/download?access_token=` | Downloads a file with its access token, to configure as `downloadurl`. |
| `GET /wopi/iop/list`, `GET /wopi/iop/conflicts` | The open files and the lock conflicts, for operators. Need the iopsecret. |
| `GET /wopi/iop/test` | A token for the Microsoft WOPI validator. Needs the iopsecret. |
| `/wopi/files/<id>`, `/wopi/files/<id>/contents` | The WOPI protocol, for the editors. |
| `/wopi/bridge/<docid>`, `GET /wopi/bridge/list` | The bridge: the save webhook of the bridged apps, and the open documents for operators. |
| `GET /metrics` | Prometheus metrics, including `open_by_app`. |
| `GET /wopi` | An information page, with the health of the storage. |


## Storage

The `local` storage maps the files of a folder (`[local] storagehomepath`). It keeps the WOPI metadata (locks, save
times, user info) in a hidden `.wopimeta` folder, and runs on Linux, Windows and macOS. It treats all users as the same
one, so it is meant for development and tests.

For production, implement `IStorage` over your storage and embed the server:

```cs
var configuration = WopiConfig.Load(defaultsPath: null, configPath: "/etc/wopi/wopiserver.conf");
var app = await WopiApp.Build(
    configuration,
    new()
    {
        Storage = new MyStorage()
    });
await app.RunAsync();
```

`IStorage` covers stat, extended attributes, locks, and reading, writing, renaming and removing files. Failures are
`StorageException`s whose message holds `Common.EnoentMsg`, `Common.ExclError` or `Common.AccessError` when they apply:
the server tells a missing file, a lock conflict and a permission error apart by them. `StorageTests.cs`, the suite of
the local storage, is a starting point to test an implementation.

`WopiAppOptions` also takes a `TimeProvider`, a `TextWriter` for the log, the `HttpMessageHandler` of the outgoing
requests, and a callback to configure the `WebApplicationBuilder`.


## Proof keys

WOPI clients sign their requests with a key they publish in their discovery. To check those signatures, list the
discovery URLs of the trusted clients:

```ini
[security]
proofkeydiscoveryurls = https://your-office/hosting/discovery https://your-collabora/hosting/discovery
```

Signed requests that match none of the keys are refused with a 500, as the Microsoft WOPI validator expects. The keys
are cached for an hour and refetched when a proof does not match, so key rotation is followed.

Unsigned requests are still accepted, unless `proofkeyrequired = True`. That setting also stops advertising `FileUrl`,
which would let a client bypass GetFile, and it does not work with the bridge, whose calls to this server are not
signed.


## Bridge

Set `codimdurl` or `etherpadurl` in `[general]`, with `codimdinturl` or `etherpadinturl` when the server reaches the
app at another URL, and put the API key in `/var/run/secrets/codimd_apikey` or `etherpad_apikey`. The apps can also be
registered on their first open, with `appname=CodiMD` or `Etherpad` in `/wopi/iop/openinapp`.

The bridge pushes the document to the app on open, locks it, and saves it back when the app calls its webhook, every
`saveinterval` seconds at most and on close. Markdown with pictures is saved as a `.zmd` zip bundle.


## Logs and metrics

The log is one JSON object per line, with `time`, `host`, `level`, `module`, `msg`, and the values of the message:

```json
{"time":"2026-10-18T16:43:51.940","host":"XPS","level":"INFO","process":"wopiserver","module":"wopiutils","msg":"No lock found","lockop":"Genaccesstoken","user":"0:0","filename":"/a.docx","token":"NA"}
```

Access tokens are logged by their last 20 characters, except tokens that fail validation, which are logged in full to
help debugging. The CheckFileInfo URLs are redacted unless `loglevel = Debug`.
`/metrics` exposes the request metrics of prometheus-net and the `open_by_app` counter.


## Development

```pwsh
dotnet build src
dotnet run --project src/WopiServer.Tests --no-build
```

The test project is a TUnit executable: run it, rather than `dotnet test`. It holds:

* `StorageTests`: the storage interface suite of the Python server, against the local storage.
* `SpecTests`: conformance to the server side of [MS-WOPI], each test citing its section.
* `EdgeCaseTests`: protocol edge cases ported from [WopiHost](https://github.com/petrsvihlik/WopiHost).
* `ProofKeyTests`: the proof key validation, including key rotation.
* `WopiTests`, `BridgeTests`, `UtilsTests`: the protocol end to end, the bridge against a stubbed CodiMD, and the
  helpers.

CI builds and runs the tests on Linux and Windows.


## Credits and license

The Python CS3 WOPI server is by Giuseppe Lo Presti (CERN) and its contributors. Like it, this server is under the
Apache 2.0 license, see [license.txt](license.txt).

[MS-WOPI]: https://learn.microsoft.com/en-us/openspecs/office_protocols/ms-wopi/
