<#
.SYNOPSIS
Creates the Hyper-V VM of the SharePoint lab, with Windows Server 2022 installed and the SQL Server and SharePoint media
in its DVD drives.

.DESCRIPTION
Applies Windows Server 2022 Datacenter (Desktop Experience) from the evaluation ISO to a new VHDX, with unattend.xml
for a first boot without questions, then creates a generation 2 VM on the Default Switch (NAT to the internet, and a
<Name>.mshome.net name the host resolves) and starts it. Run elevated (mounting the VHDX needs it), from Windows
PowerShell or PowerShell 7. The Administrator password is asked for: it is the password of the lab's Administrator,
then of LAB\Administrator. See docs/sharepoint-lab.md.

.EXAMPLE
src/sharepoint-lab/New-LabVm.ps1 -Media D:\SharePointLab\media -Path D:\SharePointLab
#>
#Requires -RunAsAdministrator
param(
    # the folder of Get-LabMedia.ps1
    [Parameter(Mandatory)]
    [string] $Media,
    # where the VM and its disk go: about 100 GB once SharePoint is installed
    [Parameter(Mandatory)]
    [string] $Path,
    [string] $Name = 'SPSE',
    [int] $MemoryGB = 16,
    [int] $Processors = 4,
    [int] $DiskGB = 150,
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator')
)

$ErrorActionPreference = 'Stop'

if (-not [Environment]::Is64BitProcess) {
    throw 'Run this from a 64-bit PowerShell, not Windows PowerShell (x86)'
}

if (Get-VM -Name $Name -ErrorAction SilentlyContinue) {
    throw "A VM named $Name exists already"
}

$windowsIso = Join-Path $Media 'WindowsServer2022.iso'
$sharePointIso = Join-Path $Media 'OfficeServer.iso'
$sqlIso = Get-ChildItem (Join-Path $Media 'sql') -Filter *.iso | Select-Object -First 1
foreach ($file in $windowsIso, $sharePointIso) {
    if (-not (Test-Path $file)) { throw "Missing ${file}: run Get-LabMedia.ps1 first" }
}
if (-not $sqlIso) { throw 'Missing the SQL Server ISO: run Get-LabMedia.ps1 first' }

New-Item -ItemType Directory -Force $Path | Out-Null
$vhd = Join-Path $Path "$Name.vhdx"
# the disk of a run that failed before making the VM
if (Test-Path $vhd) {
    Write-Host "Removing $vhd, left by a run that failed"
    Dismount-VHD -Path $vhd -ErrorAction SilentlyContinue
    Remove-Item $vhd
}

# the Windows image: Datacenter with the Desktop Experience
$mounted = Mount-DiskImage $windowsIso -PassThru
try {
    $isoDrive = ($mounted | Get-Volume).DriveLetter
    $wim = "${isoDrive}:\sources\install.wim"
    $image = Get-WindowsImage -ImagePath $wim |
        Where-Object { $_.ImageName -like '*Datacenter*' -and $_.ImageName -like '*Desktop*' } |
        Select-Object -First 1
    if (-not $image) { throw "No Datacenter (Desktop Experience) image in $wim" }
    Write-Host "Applying $($image.ImageName) to $vhd"

    # a GPT disk: EFI system partition, MSR, Windows
    New-VHD -Path $vhd -SizeBytes ($DiskGB * 1GB) -Dynamic | Out-Null
    $disk = Mount-VHD -Path $vhd -Passthru | Get-Disk
    try {
        Initialize-Disk -Number $disk.Number -PartitionStyle GPT
        $system = New-Partition -DiskNumber $disk.Number -Size 260MB -GptType '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}' -AssignDriveLetter
        Format-Volume -Partition $system -FileSystem FAT32 -NewFileSystemLabel System -Confirm:$false | Out-Null
        New-Partition -DiskNumber $disk.Number -Size 16MB -GptType '{e3c9e316-0b5c-4db8-817d-f92df00215ae}' | Out-Null
        $windows = New-Partition -DiskNumber $disk.Number -UseMaximumSize -AssignDriveLetter
        Format-Volume -Partition $windows -FileSystem NTFS -NewFileSystemLabel Windows -Confirm:$false | Out-Null
        $systemDrive = (Get-Partition -DiskNumber $disk.Number -PartitionNumber $system.PartitionNumber).DriveLetter
        $windowsDrive = (Get-Partition -DiskNumber $disk.Number -PartitionNumber $windows.PartitionNumber).DriveLetter

        Expand-WindowsImage -ImagePath $wim -Index $image.ImageIndex -ApplyPath "${windowsDrive}:\" | Out-Null
        # the boot files, with the bcdboot of the image rather than the host's, which can be of another version
        $bcdboot = "${windowsDrive}:\Windows\System32\bcdboot.exe"
        $output = & $bcdboot "${windowsDrive}:\Windows" /s "${systemDrive}:" /f UEFI 2>&1
        if ($LASTEXITCODE -ne 0) { throw "bcdboot failed: $LASTEXITCODE $output" }

        # the answers of the first boot, with the password escaped for XML
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($AdministratorPassword))
        $unattend = (Get-Content (Join-Path $PSScriptRoot 'unattend.xml') -Raw).
            Replace('{ComputerName}', $Name).
            Replace('{AdministratorPassword}', [Security.SecurityElement]::Escape($plain))
        $plain = $null
        $panther = "${windowsDrive}:\Windows\Panther"
        New-Item -ItemType Directory -Force $panther | Out-Null
        Set-Content -Path (Join-Path $panther 'unattend.xml') -Value $unattend -Encoding UTF8
    }
    finally {
        Dismount-VHD -Path $vhd
    }
}
finally {
    Dismount-DiskImage $windowsIso | Out-Null
}

# the VM, with the SQL Server and SharePoint media in its DVD drives
$vm = New-VM -Name $Name -Generation 2 -MemoryStartupBytes ($MemoryGB * 1GB) -VHDPath $vhd -SwitchName 'Default Switch' -Path $Path
Set-VM -VM $vm -ProcessorCount $Processors -AutomaticCheckpointsEnabled $false -AutomaticStopAction ShutDown
# dynamic memory, so that the VM starts beside Docker or Visual Studio, and takes up to $MemoryGB when the host has it
Set-VMMemory -VM $vm -DynamicMemoryEnabled $true -StartupBytes 8GB -MinimumBytes 4GB -MaximumBytes ($MemoryGB * 1GB)
Enable-VMIntegrationService -VM $vm -Name 'Guest Service Interface'
Add-VMDvdDrive -VM $vm -Path $sqlIso.FullName
Add-VMDvdDrive -VM $vm -Path $sharePointIso
Set-VMFirmware -VM $vm -FirstBootDevice (Get-VMHardDiskDrive -VM $vm)
Start-VM -VM $vm
Write-Host "$Name is starting: run Install-LabFarm.ps1 next"
