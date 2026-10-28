<#
.SYNOPSIS
Installs the module of the OfficeHost shadow in the lab's SharePoint, or removes it with -Remove.

.DESCRIPTION
Builds src/OfficeHost.SharePoint (Release), copies the module and its install script
(src/OfficeHost.SharePoint/Install-OfficeHostShadow.ps1) to the VM, and runs the script there through PowerShell Direct:
a farm solution deployed to the GAC, then the module in the web.config of http://<Name>.mshome.net. By default it
records to C:\ReplayRecording in the VM, for the replay (Save-ReplayRecording.ps1 moves the recording to the host);
with -Sink and -ApiKey it sends to a shadow instead. Removes the recorder of earlier versions of the lab (a module in the
bin folder) first. Run elevated. See docs/replay.md and docs/shadow.md.

.EXAMPLE
src/sharepoint-lab/Install-ReplayRecorder.ps1

.EXAMPLE
src/sharepoint-lab/Install-ReplayRecorder.ps1 -Sink https://172.25.48.1:5001/officehost-shadow/exchanges -ApiKey <key>

.EXAMPLE
src/sharepoint-lab/Install-ReplayRecorder.ps1 -Remove
#>
#Requires -RunAsAdministrator
param(
    [string] $Name = 'SPSE',
    [string] $NetBios = 'LAB',
    # a folder of the VM to record to, or the ingest of a shadow
    [string] $Sink = 'C:\ReplayRecording',
    [string] $ApiKey,
    [int] $SamplePercent = 100,
    [switch] $Remove,
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator')
)

$ErrorActionPreference = 'Stop'
$credential = New-Object PSCredential("$NetBios\Administrator", $AdministratorPassword)
$url = "http://$($Name.ToLowerInvariant()).mshome.net"
$project = Join-Path $PSScriptRoot '..\OfficeHost.SharePoint'

if (-not $Remove) {
    dotnet build $project --configuration Release | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw 'The module did not build'
    }
}

$session = New-PSSession -VMName $Name -Credential $credential
try {
    Invoke-Command -Session $session { New-Item -ItemType Directory -Force C:\OfficeHostShadow | Out-Null }
    Copy-Item (Join-Path $project 'Install-OfficeHostShadow.ps1') -Destination C:\OfficeHostShadow\ -ToSession $session
    if (-not $Remove) {
        Copy-Item (Join-Path $project 'bin\Release\net462\OfficeHost.SharePoint.dll') -Destination C:\OfficeHostShadow\ -ToSession $session
    }

    Invoke-Command -Session $session -ArgumentList $url, $Sink, $ApiKey, $SamplePercent, $Remove.IsPresent -ScriptBlock {
        param([string] $Url, [string] $Sink, [string] $ApiKey, [int] $SamplePercent, [bool] $Remove)
        $ErrorActionPreference = 'Stop'
        Import-Module SharePointServer -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if (-not (Get-Command Get-SPWebApplication -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.SharePoint.PowerShell
        }

        # the recorder of earlier versions of the lab: a module in the bin folder, owner ReplayRecorder
        $webApp = Get-SPWebApplication $Url
        $old = @($webApp.WebConfigModifications | Where-Object Owner -eq 'ReplayRecorder')
        if ($old.Count -gt 0) {
            foreach ($modification in $old) {
                $webApp.WebConfigModifications.Remove($modification) | Out-Null
            }

            $webApp.Update()
            $webApp.Parent.ApplyWebConfigModifications()
        }

        Remove-Item (Join-Path $webApp.IisSettings['Default'].Path.FullName 'bin\ReplayRecorder.dll') -ErrorAction SilentlyContinue

        if ($Remove) {
            & C:\OfficeHostShadow\Install-OfficeHostShadow.ps1 -Remove
            return
        }

        New-Item -ItemType Directory -Force C:\ReplayRecording | Out-Null
        $arguments = @{
            WebApplication = $Url
            Sink = $Sink
            SamplePercent = $SamplePercent
        }
        if ($ApiKey) {
            $arguments.ApiKey = $ApiKey
        }

        & C:\OfficeHostShadow\Install-OfficeHostShadow.ps1 @arguments
    }
}
finally {
    Remove-PSSession $session
}
