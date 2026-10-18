# WopiServer

A C# (ASP.NET Core, .NET 10) port of the [CS3 WOPI server](https://github.com/cs3org/wopiserver), the vendor-neutral
application gateway of the ScienceMesh IOP. It implements the Web-application Open Platform Interface
([WOPI](https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/)), so EFSS storages can
integrate Office Online platforms such as Microsoft 365 for the web, Collabora Online and ONLYOFFICE, plus a bridge
for CodiMD and Etherpad.

The original Python implementation is by Giuseppe Lo Presti (CERN) and contributors, under the Apache 2.0 license
(see `license.txt`); this port keeps that license.


## What is ported

* **Core WOPI protocol**: CheckFileInfo, GetFile, PutFile (including editnew), Lock, RefreshLock, UnlockAndRelock,
  GetLock, Unlock, PutRelative, RenameFile, DeleteFile, PutUserInfo. Including the interoperability with Microsoft
  Office and LibreOffice desktop lock files, external modification detection with conflict copies, the recovery
  space, the WOPI proxy and the business flow / CSPP+ settings.
* **IOP endpoints**: `/wopi/iop/openinapp`, `/wopi/iop/download`, `/wopi/iop/list`, `/wopi/iop/conflicts`,
  `/wopi/iop/test`, plus the index page and Prometheus metrics at `/metrics`.
* **Bridge**: `/wopi/bridge/<docid>` and `/wopi/bridge/list`, with the CodiMD and Etherpad plugins and the
  asynchronous save loop.
* **Storage**: the `local` storage interface, plus `IStorage` to plug in others.

Not ported:

* The `xroot` storage interface: it needs the XRootD client library, which has no .NET binding.
* The `cs3` storage interface: it needs the CS3 APIs gRPC client, which has no .NET package.
* The `tools/` scripts.


## Compatibility

* The configuration is the same INI `wopiserver.conf` as the Python server. `src/WopiServer/wopiserver.defaults.conf`
  is shipped next to the executable and loaded first, then `/etc/wopi/wopiserver.defaults.conf` when present, then
  `/etc/wopi/wopiserver.conf` (or the file given with `--config`), which is watched for changes.
* Access tokens are the same HS256 JWTs, with the same claims: given the same `wopisecret`, tokens are interchangeable
  between the two implementations. Locks are stored in the same Reva-compatible format.
* Logs are JSON lines with the same keys (`time`, `host`, `level`, `process`, `module`, `msg`, and the message values).
* The local storage keeps xattrs and locks in a hidden `.wopimeta` folder at the storage root instead of Linux user
  extended attributes, so it runs on Windows and macOS too. Its file ids are derived from the path, as there is no
  portable inode API. Like the original, it is meant for development and tests only.
* `internalserver` (flask/waitress) does not apply: the server always runs on Kestrel.


## Behaviour differences

Fixes of defects found in the Python code while porting:

* PutRelative works. In Python it always returned 501, as the token's view mode and user type strings were compared
  with enum members. Its suggested-extension target also lacked a separator (`/dirname` + `file.pdf`), its
  `HostEditUrl`/`HostViewUrl` lacked the `WOPISrc=` key, and the URL templates failed on a missing key.
* RenameFile with a requested name containing a dot renamed the file to an empty name.
* `/wopi/iop/test` and PutRelative unpacked 3 of the 4 values returned when generating a token.
* PutUserInfo passed the lock dictionary where an encoded lock id was expected.
* The local storage read xattrs keys with `str.strip('user.')`, which also stripped the trailing `e` of
  `iop.wopi.lastwritetime`, so external modifications were always detected.
* The recovery space got an empty file after a failed streamed save: request bodies are now buffered so they can be
  stored for recovery.
* CodiMD attachment collisions crashed on `bytes(str)`; the bridge also failed to find a plugin registered by an
  open request with a mixed case app name.
* `Content-Disposition` used `+` for spaces; it is now RFC 5987 percent-encoded.
* The local storage rejects paths escaping the storage root.

Changes for conformance to [MS-WOPI] (see `SpecTests.cs`, where each test cites its section):

* GetLock on an unlocked file returns 200 with an empty `X-WOPI-Lock` (the Python server returns 404, which the spec
  reserves for an unknown file). GetLock, GetFile and Unlock on an unknown file return 404.
* UnlockAndRelock (`LOCK` with `X-WOPI-OldLock`) on an unlocked file is a 409 instead of taking the new lock.
* WOPI responses include `X-WOPI-ServerVersion` and `X-WOPI-MachineName`, 500 responses `X-WOPI-ServerError`, and lock
  conflicts with desktop Office/LibreOffice locks `X-WOPI-LockedByOtherInterface: true`.
* CheckFileInfo no longer sends `BreadcrumbBrandName`/`BreadcrumbBrandUrl` as null, nor an empty `FileExtension`.
* PutRelativeFile makes suggested names legal, rejects illegal relative targets with 400 (they could escape the folder),
  and returns `X-WOPI-ValidRelativeTarget` with a 409.
* RenameFile returns 501 unless `enablerename` is set, 400 with `X-WOPI-InvalidFileNameError` for an illegal name, and
  no longer overwrites an existing file.
* PutFile with a read-only token is a 401: the Python server let it create a file in place of an empty one.
* The local storage keeps file versions from repeating when two saves fall within the file system clock resolution.

Following the Microsoft WOPI documentation, from edge cases ported from [WopiHost](https://github.com/petrsvihlik/WopiHost)
(`EdgeCaseTests.cs`):

* Lock, RefreshLock and UnlockAndRelock with a missing, empty or longer than 1024 characters `X-WOPI-Lock` are a 400.
* PutFile without `X-WOPI-Lock` on a locked file is a 409 giving the lock, even for an empty file; PutFile on a file
  deleted behind the server's back is a 404 instead of recreating it.
* GetFile returns 412 when the file is larger than `X-WOPI-MaxExpectedSize`.
* PutUserInfo rejects more than 1024 characters (400), and an unknown file is a 404.
* File names cannot hold `<>:"|?*` or end with a dot or a space, whatever the OS: suggested names are scrubbed, relative
  targets and rename names are a 400, with `X-WOPI-ValidRelativeTarget` for PutRelativeFile. On Windows, `a:b.docx`
  used to be written to an NTFS alternate data stream of a file `a`.

Not in the Python implementation:

* Proof keys: with `[security] proofkeydiscoveryurls`, the `X-WOPI-Proof`/`X-WOPI-ProofOld` signatures of WOPI requests
  are checked against the keys the WOPI clients publish in their discovery, including key rotation
  (`ProofKeyTests.cs`). Signed requests that do not match are refused; with `proofkeyrequired = True`, so are unsigned
  ones, which rules out the bridge. See `wopiserver.defaults.conf`.

Two spec requirements are not met, and their tests are skipped with the reason:

* CheckFileInfo omits the CSPP+ properties the spec lists as required (`OfficeCollaborationServiceEndpointUrl`,
  `OpenInClientCommandUrl`, `OpenInClientPostMessage`, `UserPrincipalName`), as the Python server does: some of them
  break regular Office for the web sessions.
* After RenameFile, the tokens issued before point to the old path, so the session cannot go on with the same id.
  This is why `enablerename` is off by default, as in the Python server.


## Build, test and run

```pwsh
dotnet build src
dotnet run --project src/WopiServer.Tests --no-build
dotnet run --project src/WopiServer -- --config /path/to/wopiserver.conf
```

The tests include a port of the Python storage interface test suite (`test/test_storageiface.py`), unit tests of the
helpers, a conformance suite for the server side of [MS-WOPI] (`SpecTests.cs`), edge cases ported from WopiHost
(`EdgeCaseTests.cs`), end-to-end tests of the WOPI protocol on
an in-memory test server, and a CodiMD bridge round trip against a stub app.

To run the server for development, write the two secret files and a configuration like the following, where the
paths are adapted to your machine:

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
