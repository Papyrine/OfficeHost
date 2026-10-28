<#
.SYNOPSIS
Lists the cell storage requests of Office that the lab's SharePoint answered neither 200 nor 401, from its IIS logs.

.DESCRIPTION
Reads the IIS logs of the lab's site in the VM, through PowerShell Direct, after flushing IIS's log buffer. For each
day: the cell storage requests of Word, Excel and PowerPoint answered otherwise than 200 or 401 (the handshake of NTLM),
with their Windows error (64: the client dropped the connection) and time, then a count of all of them by status and
client. Signed in, SharePoint Server SE logs `cellstorage.svc` as `cobalt.ashx/cellstorageservice`. With -Around, the
raw log lines of the 10 seconds from a time (UTC) instead. Run elevated. See docs/replay.md.

.EXAMPLE
src/sharepoint-lab/Find-CellStorageErrors.ps1

.EXAMPLE
src/sharepoint-lab/Find-CellStorageErrors.ps1 -Around '2026-10-28 06:26:30'
#>
#Requires -RunAsAdministrator
param(
    [string] $Name = 'SPSE',
    [string] $NetBios = 'LAB',
    [string] $Site = 'Lab',
    # the raw log lines of the 10 seconds from this UTC time instead, as '2026-10-28 06:26:30'
    [string] $Around,
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator')
)

$ErrorActionPreference = 'Stop'
$credential = New-Object PSCredential("$NetBios\Administrator", $AdministratorPassword)
Invoke-Command -VMName $Name -Credential $credential -ArgumentList $Site, $Around -ScriptBlock {
    param([string] $Site, [string] $Around)
    Import-Module WebAdministration
    $website = Get-Website -Name $Site
    $folder = Join-Path ([Environment]::ExpandEnvironmentVariables($website.logFile.directory)) "W3SVC$($website.id)"
    "site $Site, id $($website.id), logging enabled $($website.logFile.enabled), folder $folder"
    # the entries IIS still holds in memory
    netsh http flush logbuffer | Out-Null
    $logs = @(Get-ChildItem $folder -Filter *.log -ErrorAction SilentlyContinue | Sort-Object Name)
    "logs: $($logs.Count) files, $($logs[0].Name) to $($logs[-1].Name)"
    if ($Around) {
        $from = [DateTime]::Parse($Around)
        $log = Join-Path $folder "u_ex$($from.ToString('yyMMdd')).log"
        Get-Content $log | Where-Object { $_.StartsWith('#Fields:') } | Select-Object -First 1
        foreach ($line in Get-Content $log) {
            if ($line.StartsWith('#')) {
                continue
            }

            $at = [DateTime]::Parse($line.Substring(0, 19))
            if ($at -ge $from -and $at -lt $from.AddSeconds(10)) {
                $line
            }
        }

        return
    }

    foreach ($log in $logs) {
        $fields = $null
        $statuses = @{}
        foreach ($line in Get-Content $log.FullName) {
            if ($line.StartsWith('#Fields:')) {
                $fields = $line.Substring(9).Split(' ')
                continue
            }

            if ($line.StartsWith('#') -or -not $fields -or $line -notmatch 'cellstorage\.svc|cobalt\.ashx') {
                continue
            }

            $values = $line.Split(' ')
            $entry = @{}
            for ($index = 0; $index -lt $fields.Count; $index++) {
                $entry[$fields[$index]] = $values[$index]
            }

            $agent = "$($entry['cs(User-Agent)'])"
            $key = "$($entry['sc-status']) $($agent -replace '\+', ' ' -replace '^(Microsoft Office \w+).*', '$1')"
            $statuses[$key] = 1 + $statuses[$key]
            if ($agent -match 'PowerPoint|Word|Excel' -and $entry['sc-status'] -notin '401', '200') {
                "  $($entry['date']) $($entry['time']) $($entry['sc-status']) win32 $($entry['sc-win32-status']) $($entry['time-taken']) ms $($entry['cs-username']) $agent"
            }
        }

        "$($log.Name): cell storage requests by status and client: $(($statuses.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) x$($_.Value)" }) -join '; ')"
    }
}
