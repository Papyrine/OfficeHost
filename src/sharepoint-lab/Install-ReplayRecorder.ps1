<#
.SYNOPSIS
Installs the replay recorder in the lab's SharePoint, or removes it with -Remove.

.DESCRIPTION
Copies ReplayRecorder.cs to the VM and, there, through PowerShell Direct: compiles it with the C# compiler of the .NET
Framework into the bin folder of the web application, and adds it to the modules of the web application, with the
folder it records to (C:\ReplayRecording in the VM), as SharePoint web.config modifications. From then on, the cell
storage and shared access requests of the site are recorded with their responses, for the replay of OfficeHost;
Save-ReplayRecording.ps1 moves them to the host. Run elevated. See docs/replay.md.

.EXAMPLE
src/sharepoint-lab/Install-ReplayRecorder.ps1

.EXAMPLE
src/sharepoint-lab/Install-ReplayRecorder.ps1 -Remove
#>
#Requires -RunAsAdministrator
param(
    [string] $Name = 'SPSE',
    [string] $NetBios = 'LAB',
    # takes the recorder out of the web application
    [switch] $Remove,
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator')
)

$ErrorActionPreference = 'Stop'
$credential = New-Object PSCredential("$NetBios\Administrator", $AdministratorPassword)
$url = "http://$($Name.ToLowerInvariant()).mshome.net"

$session = New-PSSession -VMName $Name -Credential $credential
try {
    if (-not $Remove) {
        Invoke-Command -Session $session { New-Item -ItemType Directory -Force C:\ReplayRecorder | Out-Null }
        Copy-Item (Join-Path $PSScriptRoot ReplayRecorder.cs) -Destination C:\ReplayRecorder\ -ToSession $session
    }

    Invoke-Command -Session $session -ArgumentList $url, $Remove.IsPresent -ScriptBlock {
        param([string] $Url, [bool] $Remove)
        $ErrorActionPreference = 'Stop'
        Import-Module SharePointServer -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if (-not (Get-Command Get-SPWebApplication -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.SharePoint.PowerShell
        }

        $owner = 'ReplayRecorder'
        $webApp = Get-SPWebApplication $Url
        $site = $webApp.IisSettings['Default'].Path.FullName
        $dll = Join-Path $site bin\ReplayRecorder.dll

        # what a previous install added
        foreach ($modification in @($webApp.WebConfigModifications | Where-Object Owner -eq $owner)) {
            $webApp.WebConfigModifications.Remove($modification) | Out-Null
        }
        $webApp.Update()
        $webApp.Parent.ApplyWebConfigModifications()
        if ($Remove) {
            Remove-Item $dll -ErrorAction SilentlyContinue
            Write-Host "Removed the recorder from $Url"
            return
        }

        # an assembly of the bin folder runs with the trust level of the web application: unsigned, out of the GAC, the
        # recorder needs full trust to write its files
        $config = [xml] (Get-Content (Join-Path $site web.config))
        $trust = $config.SelectSingleNode('/configuration/system.web/trust/@level').Value
        if ($trust -ne 'Full') {
            throw "The web application of $Url has the trust level $trust, not Full: the recorder in its bin folder could not write"
        }

        # the recorder reads the body before the service: WCF reads it again only in the ASP.NET compatibility mode
        if (-not $config.SelectSingleNode("//serviceHostingEnvironment[@aspNetCompatibilityEnabled='true']")) {
            Write-Warning "The web.config of $Url does not enable the ASP.NET compatibility of WCF: the cell storage service may fail to read the requests"
        }

        $csc = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) csc.exe
        & $csc /nologo /target:library /optimize "/out:$dll" /reference:System.Web.dll /reference:System.Configuration.dll C:\ReplayRecorder\ReplayRecorder.cs
        if ($LASTEXITCODE -ne 0) {
            throw 'ReplayRecorder.cs did not compile'
        }

        $folder = 'C:\ReplayRecording'
        New-Item -ItemType Directory -Force $folder | Out-Null
        $module = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
        $module.Path = 'configuration/system.webServer/modules'
        $module.Name = "add[@name='ReplayRecorder']"
        $module.Value = '<add name="ReplayRecorder" type="ReplayRecorder, ReplayRecorder" />'
        $setting = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
        $setting.Path = 'configuration/appSettings'
        $setting.Name = "add[@key='ReplayRecorder.Folder']"
        $setting.Value = "<add key=`"ReplayRecorder.Folder`" value=`"$folder`" />"
        foreach ($modification in $module, $setting) {
            $modification.Owner = $owner
            $modification.Sequence = 0
            $modification.Type = [Microsoft.SharePoint.Administration.SPWebConfigModification+SPWebConfigModificationType]::EnsureChildNode
            $webApp.WebConfigModifications.Add($modification)
        }
        $webApp.Update()
        $webApp.Parent.ApplyWebConfigModifications()

        # what the web application has now, as IIS reads it
        if (-not (Test-Path $dll)) {
            throw "No $dll after the install"
        }

        if (-not (Select-String -Path (Join-Path $site web.config) -Pattern 'ReplayRecorder' -Quiet)) {
            throw "The web.config of $Url does not have the recorder: SharePoint did not apply the modifications"
        }

        Import-Module WebAdministration
        $iisSite = $webApp.IisSettings['Default'].ServerComment
        $modules = Get-WebConfiguration -PSPath "IIS:\Sites\$iisSite\_vti_bin" -Filter system.webServer/modules/add | ForEach-Object name
        if ($modules -notcontains 'ReplayRecorder') {
            throw "IIS does not load the recorder for $Url/_vti_bin"
        }

        Write-Host "The recorder records $Url to $folder in the VM"
    }
}
finally {
    Remove-PSSession $session
}
