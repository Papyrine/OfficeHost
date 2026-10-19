# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A library for ASP.NET Core that implements the host side of WOPI ([MS-WOPI]): `AddWopi` registers the services,
`MapWopi` maps the `/wopi/files` endpoints, and the application supplies the storage. It started as a port of the
Python CS3 WOPI server, but is not compatible with it and does not follow it: design on the merits of WOPI and ASP.NET
Core, with no CS3, Reva or container specifics.

## Commands

- Build: `dotnet build src` (warnings are errors, code style is enforced in the build).
- Tests: `dotnet run --project src/WopiServer.Tests --no-build` (TUnit, do not use `dotnet test`).
  One test: `... -- --treenode-filter '/*/*/SpecTests/Lock_UnlockedFile'`.
- Coverage: see the readme (`src/coverage.config`, `src/coverage-summary.ps1`).
- WOPI validator: `pwsh src/wopi-validator.ps1` (runs `ValidatorTests` of `src/WopiServer.E2ETests`).
- Editor tests (Docker, Playwright): `src/WopiServer.E2ETests`, `EditorTests`, see the readme.
- Sample: `dotnet run --project src/WopiServer.Sample` (an application hosting the library, tested by `SampleTests`).
- Readme snippets: the Usage code of `readme.md` comes from the `begin-snippet` regions of the sample. Edit the code, not
  the readme, then run `mdsnippets .` (the `on-push-do-docs.yml` workflow also does on push).
- CI: `.github/workflows/build.yml` (build, tests with coverage on Ubuntu and Windows, validator), `e2e.yml` (editors,
  nightly).

## Layout

- `WopiExtensions.cs`: `AddWopi` and `MapWopi`, with the endpoint filter adding the X-WOPI headers of [MS-WOPI] 2.2.1
  and turning unexpected exceptions into 500s.
- `WopiHandler.cs`: the WOPI operations, over `IWopiStorage`, `IWopiLockStore` and `IWopiTokens`. Locks go through the
  store's `CompareExchange`, so each lock operation is atomic.
- `WopiOptions.cs`, `WopiAccess.cs` (the access a token grants, `IWopiTokens` and the built-in signed tokens on
  `Jwt.cs`), `WopiLauncher.cs` (action URLs from the discovery), `WopiDiscovery.cs` and `ProofKeys.cs`.
- `IWopiStorage.cs` (with `IWopiUserInfoStore`), `FileSystemStorage.cs`, `IWopiLockStore.cs` (with `MemoryLockStore`),
  `FileNames.cs` (legal names, UTF-7 names of PutRelativeFile and RenameFile).

## Conventions

- ProjectDefaults (nuget) overwrites the root `.editorconfig` on build: don't hand-edit it. No namespaces, implicit
  accessibility modifiers, central package management in `src/Directory.Packages.props`.
- The public API is what the application uses: keep implementation types internal (the tests see them through
  `InternalsVisibleTo`).
- Responses follow [MS-WOPI] first, then the Microsoft WOPI documentation; the WOPI validator decides where they differ
  (as the 500 for an invalid X-WOPI-TimeStamp or proof).

[MS-WOPI]: https://learn.microsoft.com/en-us/openspecs/office_protocols/ms-wopi/
