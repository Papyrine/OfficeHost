# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

OfficeHost (formerly WopiServer) is a library for ASP.NET Core that implements the host side of the Office document
protocols: WOPI ([MS-WOPI]) for Office for the web, Collabora and ONLYOFFICE, and, in progress on the `fsshttp` branch,
MS-FSSHTTP for desktop Office coauthoring (docs/fsshttp-spike.md, docs/fsshttp-notes.md). `AddOfficeHost` registers the
services, `MapOfficeHost` maps the endpoints, and the application supplies the storage. Everything is in the one
`OfficeHost` project and package. The WOPI part started as a port of the Python CS3 WOPI server, but is not compatible
with it and does not follow it: design on the merits of the protocols and ASP.NET Core, with no CS3, Reva or container
specifics.

## Commands

- Build: `dotnet build src` (warnings are errors, code style is enforced in the build).
- Tests: `dotnet run --project src/OfficeHost.Tests --no-build` (TUnit, do not use `dotnet test`).
  One test: `... -- --treenode-filter '/*/*/SpecTests/Lock_UnlockedFile'`.
- Coverage: see the readme (`src/coverage.config`, `src/coverage-summary.ps1`).
- WOPI validator: `pwsh src/wopi-validator.ps1` (runs `ValidatorTests` of `src/OfficeHost.E2ETests`).
- Editor tests (Docker, Playwright): `src/OfficeHost.E2ETests`, `EditorTests`, see the readme.
- Sample: `dotnet run --project src/OfficeHost.Sample` (an application hosting the library, tested by `SampleTests`).
- Readme snippets: the Usage code of `readme.md` comes from the `begin-snippet` regions of the sample. Edit the code, not
  the readme, then run `mdsnippets .` (the `on-push-do-docs.yml` workflow also does on push).
- CI: `.github/workflows/build.yml` (build, tests with coverage on Ubuntu and Windows, validator, and on a tag the
  nuget.org push through trusted publishing, the tag being the version), `e2e.yml` (editors, nightly).

## Layout

- `OfficeHostExtensions.cs`: `AddOfficeHost` and `MapOfficeHost`, with the endpoint filter adding the X-WOPI headers of
  [MS-WOPI] 2.2.1 and turning unexpected exceptions into 500s.
- `OfficeHostOptions.cs`, `OfficeAccess.cs` (the access a token grants, `IOfficeTokens` and the built-in signed tokens
  on `Jwt.cs`), `OfficeLauncher.cs` (the links that open a file in a WOPI client or in desktop Office), `FileNames.cs`
  (legal names, UTF-7 names of PutRelativeFile and RenameFile).
- `Storage/`: `IOfficeStorage.cs` (with `IOfficeUserInfoStore`), `FileSystemStorage.cs`, and the state of desktop
  coauthoring, opaque to its store: `IFssHttpStore.cs`, `MemoryFssHttpStore.cs`, `FileSystemFssHttpStore.cs`. `Lock/`:
  `IOfficeLockStore.cs`, `MemoryLockStore.cs`.
- `Wopi/`: [MS-WOPI]. `WopiHandler.cs`, the WOPI operations, over `IOfficeStorage`, `IOfficeLockStore` and
  `IOfficeTokens`; locks go through the store's `CompareExchange`, so each lock operation is atomic. `WopiResult.cs`,
  `WopiDiscovery.cs` and `WopiAction.cs` (the discovery of the WOPI clients), `ProofKeys.cs`, and what launching a WOPI
  client gives (`OfficeLaunch.cs`, `OfficeClientUnreachableException.cs`). Its tests are in `src/OfficeHost.Tests/Wopi`.
- `FssHttp/`: MS-FSSHTTP, internal, with its tests in `src/OfficeHost.Tests/FssHttp`. `FssHttp/Binary/` is the FSSHTTPB codec ([MS-FSSHTTPB]), pure, with no ASP.NET.
  `FssHttp/Files/` is a file stream in a cell ([MS-FSSHTTPD]): `FileChunker` (zip and simple chunking), the node
  objects, and `FileCell`, which reads the file from data elements and writes a file as the next revision.
  `FssHttp/Cells/`: `CellState`, the state of a partition of a file serialized for the `IFssHttpStore`, and `FssHttpCells`,
  which keeps the content partition in step with the file in the `IOfficeStorage`.
  `FssHttp/Service/`: the desktop sites of `MapFssHttp` (`FssHttpHandler`: OPTIONS, cellstorage.svc, sharedaccess.asmx,
  GET), SOAP and MTOM (`SoapMessage`), the sub-requests (`CellStorageService`), Query and Put Changes (`CellSync`), and
  the coauthoring tracker with its lock in the `IOfficeLockStore` (`CoauthSessions`).
- Manual test with Word: the `desktop` launch profile of the sample (HTTPS, requests logged to `request-logs` by `RequestLog.cs`), then
  `/desktop/{fileId}` (`?user=` for a second user); the steps are in docs/fsshttp-checklist.md. The simulated clients
  are `WordClient` of `src/OfficeHost.Tests/FssHttp`.

## Conventions

- ProjectDefaults (nuget) overwrites the root `.editorconfig` on build: don't hand-edit it. Implicit accessibility
  modifiers, central package management in `src/Directory.Packages.props`.
- Namespaces: the public types of the library are in `namespace OfficeHost;`, and a file with a public type holds no
  internal type. Internal types, and the sample and test projects, have no namespace: the projects import `OfficeHost`
  with a `<Using>`.
- The public API is what the application uses: keep implementation types internal (the tests see them through
  `InternalsVisibleTo`).
- Responses follow [MS-WOPI] first, then the Microsoft WOPI documentation; the WOPI validator decides where they differ
  (as the 500 for an invalid X-WOPI-TimeStamp or proof).

[MS-WOPI]: https://learn.microsoft.com/en-us/openspecs/office_protocols/ms-wopi/
[MS-FSSHTTPB]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/
[MS-FSSHTTPD]: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpd/
