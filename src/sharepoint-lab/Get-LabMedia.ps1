<#
.SYNOPSIS
Downloads the evaluation media of the SharePoint lab: Windows Server 2022, SQL Server 2022 Developer and SharePoint
Server Subscription Edition.

.DESCRIPTION
About 7.3 GB in all, from Microsoft only: the Windows Server 2022 evaluation ISO (4.7 GB, Evaluation Center),
OfficeServer.iso (1.4 GB, Download Center 103599), and the SQL Server 2022 Developer ISO (1.1 GB, the media its
bootstrapper downloads, which the bootstrapper fails to on a machine with a newer SQL Server). A file already there is
kept. Needs no elevation. See docs/sharepoint-lab.md.

.EXAMPLE
pwsh src/sharepoint-lab/Get-LabMedia.ps1 -Path D:\SharePointLab\media
#>
param(
    [Parameter(Mandatory)]
    [string] $Path
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
New-Item -ItemType Directory -Force $Path | Out-Null

$media = @(
    @{ Name = 'WindowsServer2022.iso'; Url = 'https://go.microsoft.com/fwlink/p/?LinkID=2195280&clcid=0x409&culture=en-us&country=US' },
    @{ Name = 'OfficeServer.iso'; Url = 'https://download.microsoft.com/download/3/f/5/3f5f8a7e-462b-41ff-a5b2-04bdf5821ceb/OfficeServer.iso' },
    @{ Name = 'sql\SQLServer2022-x64-ENU-Dev.iso'; Url = 'https://download.microsoft.com/download/3/8/d/38de7036-2433-4207-8eae-06e247e17b25/SQLServer2022-x64-ENU-Dev.iso' }
)

foreach ($item in $media) {
    $file = Join-Path $Path $item.Name
    if (Test-Path $file) {
        Write-Host "Have $($item.Name)"
        continue
    }

    Write-Host "Downloading $($item.Name)"
    New-Item -ItemType Directory -Force (Split-Path $file) | Out-Null
    # to a temporary name, so that an interrupted download is not taken for a whole one
    Invoke-WebRequest $item.Url -OutFile "$file.part" -UseBasicParsing
    Move-Item "$file.part" $file
}

Get-ChildItem $Path -Recurse -Filter *.iso | ForEach-Object { '{0} {1:N2} GB' -f $_.FullName, ($_.Length / 1GB) }
