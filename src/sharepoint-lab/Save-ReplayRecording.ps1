<#
.SYNOPSIS
Moves what the replay recorder recorded in the lab to a folder of the host, for the replay of OfficeHost.

.DESCRIPTION
Copies the exchanges of C:\ReplayRecording in the VM (and errors.txt, if the recorder had errors) to -Destination,
then empties the folder in the VM, so that the next recording starts there. Close Office first: a request still
running would be cut in two. Run elevated. See docs/replay.md.

.EXAMPLE
src/sharepoint-lab/Save-ReplayRecording.ps1 -Destination D:/SharePointLab/recordings/word-open-save
#>
#Requires -RunAsAdministrator
param(
    [Parameter(Mandatory)]
    [string] $Destination,
    [string] $Name = 'SPSE',
    [string] $NetBios = 'LAB',
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator')
)

$ErrorActionPreference = 'Stop'
$credential = New-Object PSCredential("$NetBios\Administrator", $AdministratorPassword)
$folder = 'C:\ReplayRecording'

$session = New-PSSession -VMName $Name -Credential $credential
try {
    $files = Invoke-Command -Session $session { Get-ChildItem $using:folder -File | Select-Object -ExpandProperty Name }
    if (-not $files) {
        throw "Nothing recorded in $folder of ${Name}: is the recorder installed (Install-ReplayRecorder.ps1)?"
    }

    New-Item -ItemType Directory -Force $Destination | Out-Null
    foreach ($file in $files) {
        Copy-Item "$folder\$file" -Destination $Destination -FromSession $session
    }

    Invoke-Command -Session $session { Get-ChildItem $using:folder -File | Remove-Item -Force }
    $exchanges = @($files | Where-Object { $_ -like '*.request.http' }).Count
    Write-Host "Saved $exchanges exchanges to $Destination"
    if ($files -contains 'errors.txt') {
        Write-Warning "The recorder had errors: see $(Join-Path $Destination errors.txt)"
    }
}
finally {
    Remove-PSSession $session
}
