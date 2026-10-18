# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A C# port of the Python CS3 WOPI server (https://github.com/cs3org/wopiserver, checked out at
`C:\Code\PythonWopiServer` on the author's machine). `docs/porting.md` lists what is and is not ported, and the
deliberate behaviour differences. When changing WOPI logic, compare with the Python source: the port follows it
function by function, keeping its comments and log messages.

## Commands

- Build: `dotnet build src` (warnings are errors, code style is enforced in the build).
- Tests: `dotnet run --project src/WopiServer.Tests --no-build` (TUnit, do not use `dotnet test`).
  One test: `... -- --treenode-filter '/*/*/WopiTests/Lock_Lifecycle'`.
- Run: `dotnet run --project src/WopiServer -- --config <wopiserver.conf>`.
- Coverage: see the readme (`src/coverage.config`, `src/coverage-summary.ps1`). WOPI validator: `pwsh src/wopi-validator.ps1`.
- CI: `.github/workflows/build.yml` runs the Release build and the tests on Ubuntu and Windows, uploading the TRX
  results, and the `.received` snapshots on failure.

## Layout

- `WopiApp.cs`: startup (the Python `Wopi.init`), `WopiEndpoints.cs`: the routes of `wopiserver.py`,
  `WopiState.cs`: settings and the open files / conflicts bookkeeping, `WopiConfig.cs`: INI config with
  `configparser` semantics.
- `Core/`: `WopiCore` (`wopi.py`), `WopiUtils` (`wopiutils.py`), `Common` (`commoniface.py`), plus Python idioms the
  protocol depends on: `PosixPath` (`os.path`, storage paths always use `/`), `UrlCodec` (`quote_plus`), `Jwt` (PyJWT
  HS256, hand-rolled as Microsoft's handlers reject short keys), `PyStrings`.
- `Storage/`: `IStorage` and `LocalStorage`. Storage errors are `StorageException`s classified by message
  (`Common.EnoentMsg`, `ExclError`, `AccessError`), as in Python.
- `Bridge/`: `Bridge` (`bridge/__init__.py`, including the save loop), `WopiClient`, `CodiMD`, `Etherpad`.
- `Logging/JsonLogProvider.cs`: JSON lines in the Python format. The `msg` is the template text before the first
  placeholder, so log templates are `"Message text {key1} {key2}"` with lowercase keys matching the Python logs.

## Conventions

- ProjectDefaults (nuget) overwrites the root `.editorconfig` on build: don't hand-edit it. No namespaces, implicit
  accessibility modifiers, central package management in `src/Directory.Packages.props`.
- The app is neither packed nor strong-named (`src/Directory.Build.props`).
- Snapshot tests use Verify: accept a change by replacing the `.verified.txt` with the reviewed `.received.txt`.
