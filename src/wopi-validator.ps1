<#
.SYNOPSIS
Runs Microsoft's WOPI validator (https://github.com/microsoft/wopi-validator-core) against the server.

.DESCRIPTION
Builds the validator at a pinned commit, starts the server (built beforehand with `dotnet build src`) over a temporary
local storage, gets a token for a .wopitest file from /wopi/iop/test, and runs each test group of the validator, on a
fresh storage each time. Fails when any test fails.

The ProofKeys group is not run: the validator's command line has no proof keys, so it sends no X-WOPI-Proof, and the
tests expecting a rejection cannot pass. ProofKeyTests.cs covers the proof keys instead.

.EXAMPLE
pwsh src/wopi-validator.ps1 -Configuration Debug
#>
param(
    [string] $Configuration = 'Release',
    [int] $Port = 18881,
    [string] $ValidatorCommit = 'b37a6c27af243be5e7c78a3bee3ee605fcc90cb6',
    [string[]] $ExcludedGroups = @('ProofKeys')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$work = Join-Path ([IO.Path]::GetTempPath()) "wopi-validator-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
$validatorSource = Join-Path $work 'validator'
$storage = Join-Path $work 'storage'
New-Item -ItemType Directory -Force $validatorSource, $storage, (Join-Path $work 'recovery') | Out-Null

Write-Host "Building the WOPI validator $ValidatorCommit"
git -C $validatorSource init --quiet
git -C $validatorSource fetch --quiet --depth 1 https://github.com/microsoft/wopi-validator-core.git $ValidatorCommit
git -C $validatorSource checkout --quiet FETCH_HEAD
dotnet build (Join-Path $validatorSource 'src/WopiValidator/WopiValidator.csproj') --configuration Release --framework net10.0 --nologo --verbosity quiet
if ($LASTEXITCODE -ne 0) { throw 'Failed to build the WOPI validator' }
$validator = Join-Path $validatorSource 'src/WopiValidator/bin/Release/net10.0/Microsoft.Office.WopiValidator.dll'

Set-Content (Join-Path $work 'wopisecret') 'validatorwopisecret'
Set-Content (Join-Path $work 'iopsecret') 'validatoriopsecret'
$config = Join-Path $work 'wopiserver.conf'
Set-Content $config @"
[general]
storagetype = local
port = $Port
wopiurl = http://localhost:$Port
loglevel = Debug
loghandler = file
logdest = $(Join-Path $work 'wopiserver.log')
[security]
wopisecretfile = $(Join-Path $work 'wopisecret')
iopsecretfile = $(Join-Path $work 'iopsecret')
[io]
recoverypath = $(Join-Path $work 'recovery')
[local]
storagehomepath = $storage
"@

function Reset-Storage {
    Get-ChildItem $storage -Force | Remove-Item -Recurse -Force
    New-Item -ItemType File (Join-Path $storage 'test.wopitest') | Out-Null
}

Reset-Storage
$server = Join-Path $root "src/WopiServer/bin/$Configuration/net10.0/wopiserver.dll"
Write-Host "Starting $server on port $Port"
$process = Start-Process dotnet -ArgumentList @($server, '--config', $config) -PassThru -NoNewWindow `
    -RedirectStandardOutput (Join-Path $work 'stdout.txt') -RedirectStandardError (Join-Path $work 'stderr.txt')

$failed = [System.Collections.Generic.List[string]]::new()
$passed = 0
try {
    $started = $false
    foreach ($attempt in 1..60) {
        try {
            Invoke-WebRequest "http://localhost:$Port/wopi" -UseBasicParsing | Out-Null
            $started = $true
            break
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }
    if (-not $started) {
        Get-Content (Join-Path $work 'stderr.txt') -ErrorAction SilentlyContinue
        throw 'The server did not start'
    }

    $test = Invoke-RestMethod "http://localhost:$Port/wopi/iop/test?filepath=/test.wopitest&usertoken=0:0" `
        -Headers @{ Authorization = 'Bearer validatoriopsecret' }
    if ($test -notmatch 'WOPI_URL=(\S+) -e WOPI_TOKEN=(\S+)') { throw "Unexpected /wopi/iop/test response: $test" }
    $wopiUrl = $Matches[1]
    $token = $Matches[2]

    $groups = ([xml] (Get-Content (Join-Path $validatorSource 'TestCases.xml') -Raw)).SelectNodes('//TestGroup') |
        ForEach-Object Name | Where-Object { $_ -notin $ExcludedGroups }
    # the validator reads its test definitions (TestCases.xml) from the current directory
    Push-Location (Split-Path $validator)
    foreach ($group in $groups) {
        Reset-Storage
        $output = dotnet $validator --wopisrc $wopiUrl --token $token --token_ttl 0 --ignore-skipped --testgroup $group 2>&1 | Out-String
        $passed += ([regex]::Matches($output, '(?m)^\s+Pass: ')).Count
        foreach ($match in [regex]::Matches($output, '(?m)^\s+Fail: (\S+)')) {
            $failed.Add($match.Groups[1].Value)
        }
        if ($LASTEXITCODE -ne 0 -and $output -notmatch 'Fail: ') {
            # the validator failed without reporting a test, for instance on an exception
            $failed.Add("$group (exit code $LASTEXITCODE)")
        }
        Write-Host $output.TrimEnd()
    }
}
finally {
    Pop-Location -ErrorAction SilentlyContinue
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host "WOPI validator: $passed passed, $($failed.Count) failed (excluded groups: $($ExcludedGroups -join ', '))"
if ($failed.Count -gt 0) {
    $failed | ForEach-Object { Write-Host "  failed: $_" }
    Write-Host "Server log: $(Join-Path $work 'wopiserver.log')"
    exit 1
}

Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
