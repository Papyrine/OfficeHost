<#
.SYNOPSIS
Runs Microsoft's WOPI validator (https://github.com/microsoft/wopi-validator-core) against the WOPI endpoints.

.DESCRIPTION
Builds the validator at a pinned commit, then runs ValidatorTests of the end-to-end test project, which hosts the
endpoints over a FileSystemStorage and runs each test group of the validator against them. Fails when any test fails.

.EXAMPLE
pwsh src/wopi-validator.ps1 -Configuration Debug
#>
param(
    [string] $Configuration = 'Release',
    [string] $ValidatorCommit = 'b37a6c27af243be5e7c78a3bee3ee605fcc90cb6'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$work = Join-Path ([IO.Path]::GetTempPath()) "wopi-validator-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
New-Item -ItemType Directory -Force $work | Out-Null

Write-Host "Building the WOPI validator $ValidatorCommit"
git -C $work init --quiet
git -C $work fetch --quiet --depth 1 https://github.com/microsoft/wopi-validator-core.git $ValidatorCommit
git -C $work checkout --quiet FETCH_HEAD
dotnet build (Join-Path $work 'src/WopiValidator/WopiValidator.csproj') --configuration Release --framework net10.0 --nologo --verbosity quiet
if ($LASTEXITCODE -ne 0) { throw 'Failed to build the WOPI validator' }

$env:WOPI_VALIDATOR = Join-Path $work 'src/WopiValidator/bin/Release/net10.0/Microsoft.Office.WopiValidator.dll'
try {
    dotnet run --project (Join-Path $root 'src/WopiServer.E2ETests') --configuration $Configuration -- --treenode-filter '/*/*/ValidatorTests/*' --output Detailed
    if ($LASTEXITCODE -ne 0) { throw 'The WOPI validator failed' }
}
finally {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}
