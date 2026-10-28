<#
.SYNOPSIS
Installs the module of the OfficeHost shadow in SharePoint Server, or removes it with -Remove.

.DESCRIPTION
Run on one server of the farm, as a farm administrator, in an elevated SharePoint Management Shell (Windows PowerShell
5.1), in a maintenance window: deploying to the GAC recycles the application pools of every server. See docs/shadow.md.

1. Builds a farm solution (WSP) of OfficeHost.SharePoint.dll with makecab, adds it to the farm and deploys it to the GAC
   of every server; an update of a deployed solution recycles instead.
2. Waits until it is deployed on every server, and only then adds the module and its settings to the web.config of the
   web applications given (SPWebConfigModification, owner OfficeHost.Shadow): a web.config that names the module before
   the assembly is in a server's GAC would fail every request of that server.
3. Creates the event log source OfficeHost.Shadow on this server (the module writes its warnings there).

-Remove takes the module out of the web.config of every web application first, waits for it, then retracts and removes
the solution.

.EXAMPLE
.\Install-OfficeHostShadow.ps1 -WebApplication https://sharepoint.contoso.com -Sink https://shadow.contoso.com/officehost-shadow/exchanges -ApiKey <key> -SamplePercent 10

.EXAMPLE
.\Install-OfficeHostShadow.ps1 -WebApplication http://spse.mshome.net -Sink C:\ReplayRecording

.EXAMPLE
.\Install-OfficeHostShadow.ps1 -Remove
#>
param(
    # the web applications to copy the traffic of
    [string[]] $WebApplication,
    # the ingest of the shadow, or a folder to record to (on each server)
    [string] $Sink,
    [string] $ApiKey,
    [ValidateRange(0, 100)]
    [int] $SamplePercent = 100,
    [int] $MaxInFlightMegabytes = 256,
    [int] $MaxExchangeMegabytes = 64,
    # path prefixes to copy, none for all
    [string[]] $Include = @(),
    # the module, by default next to the script
    [string] $Dll = (Join-Path $PSScriptRoot 'OfficeHost.SharePoint.dll'),
    [switch] $Remove,
    [int] $TimeoutMinutes = 15
)

$ErrorActionPreference = 'Stop'
$solutionName = 'officehost.sharepoint.wsp'
# fixed, so that every version of the module is the same solution
$solutionId = '7d3c1b52-3a8e-4f5c-9a61-5c2b8f0e4d17'
$owner = 'OfficeHost.Shadow'

Import-Module SharePointServer -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
if (-not (Get-Command Get-SPSolution -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

function Wait-Solution([scriptblock] $Done, [string] $What) {
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while ((Get-Date) -lt $deadline) {
        $solution = Get-SPSolution $solutionName -ErrorAction SilentlyContinue
        if (& $Done $solution) {
            return
        }

        Start-Sleep -Seconds 5
    }

    throw "The solution is not $What after $TimeoutMinutes minutes: see Central Administration, Farm Solutions"
}

function Remove-Modifications([Microsoft.SharePoint.Administration.SPWebApplication] $Application) {
    $mine = @($Application.WebConfigModifications | Where-Object Owner -eq $owner)
    foreach ($modification in $mine) {
        $Application.WebConfigModifications.Remove($modification) | Out-Null
    }

    if ($mine.Count -gt 0) {
        $Application.Update()
    }

    return $mine.Count
}

# the web.config timer job of ApplyWebConfigModifications, until it has run on every server and is gone
function Wait-WebConfig {
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while (@(Get-SPTimerJob | Where-Object Name -like 'job-webconfig-modification*').Count -gt 0) {
        if ((Get-Date) -ge $deadline) {
            throw "The web.config timer job has not run on every server after $TimeoutMinutes minutes: see Central Administration, Timer Job Status"
        }

        Start-Sleep -Seconds 5
    }
}

function Test-ModuleInWebConfig([Microsoft.SharePoint.Administration.SPWebApplication] $Application) {
    $site = $Application.IisSettings['Default']
    return Select-String -Path (Join-Path $site.Path.FullName 'web.config') -Pattern 'OfficeHost.ShadowModule' -Quiet
}

if ($Remove) {
    $changed = @()
    foreach ($application in Get-SPWebApplication) {
        if ((Remove-Modifications $application) -gt 0) {
            $changed += $application
        }
    }

    if ($changed.Count -gt 0) {
        [Microsoft.SharePoint.Administration.SPWebService]::ContentService.ApplyWebConfigModifications()
        Write-Host "Removed the module from the web.config of the web applications; waiting for every server"
        # the assembly stays in the GAC until no web.config names the module: one that did would fail every request
        Wait-WebConfig
        foreach ($application in $changed) {
            if (Test-ModuleInWebConfig $application) {
                throw "The web.config of $($application.Url) still has the module on this server: the solution is left deployed"
            }
        }
    }

    $solution = Get-SPSolution $solutionName -ErrorAction SilentlyContinue
    if ($solution) {
        if ($solution.Deployed) {
            Uninstall-SPSolution $solutionName -Confirm:$false
            Wait-Solution { param($s) -not $s.Deployed -and -not $s.JobExists } 'retracted'
        }

        Remove-SPSolution $solutionName -Confirm:$false
    }

    Write-Host 'The module of the OfficeHost shadow is removed'
    return
}

if (-not $WebApplication -or -not $Sink) {
    throw 'Pass -WebApplication and -Sink (or -Remove)'
}

if ($Sink -match '^https?://' -and -not $ApiKey) {
    throw 'Pass -ApiKey, the ApiKey of the shadow, with an HTTP sink'
}

if (-not (Test-Path $Dll)) {
    throw "No module at $Dll"
}

$assembly = [Reflection.AssemblyName]::GetAssemblyName((Resolve-Path $Dll))
if (-not $assembly.GetPublicKeyToken()) {
    throw "$Dll is not signed: the GAC takes signed assemblies only"
}

# 1. the solution: a cab of the manifest and the assembly
$work = Join-Path ([IO.Path]::GetTempPath()) "officehost-shadow-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory $work | Out-Null
try {
    Copy-Item $Dll (Join-Path $work 'OfficeHost.SharePoint.dll')
    Set-Content (Join-Path $work 'manifest.xml') -Encoding UTF8 -Value @"
<?xml version="1.0" encoding="utf-8"?>
<Solution xmlns="http://schemas.microsoft.com/sharepoint/" SolutionId="$solutionId" DeploymentServerType="WebFrontEnd" ResetWebServer="TRUE" ResetWebServerModeOnUpgrade="Recycle">
  <Assemblies>
    <Assembly Location="OfficeHost.SharePoint.dll" DeploymentTarget="GlobalAssemblyCache" />
  </Assemblies>
</Solution>
"@
    Set-Content (Join-Path $work 'solution.ddf') -Encoding ASCII -Value @"
.OPTION EXPLICIT
.Set CabinetNameTemplate=$solutionName
.Set DiskDirectoryTemplate=$work
.Set CompressionType=MSZIP
.Set UniqueFiles=on
.Set Cabinet=on
.Set MaxDiskSize=0
"$(Join-Path $work 'manifest.xml')" manifest.xml
"$(Join-Path $work 'OfficeHost.SharePoint.dll')" OfficeHost.SharePoint.dll
"@
    Push-Location $work
    try {
        & makecab.exe /F solution.ddf | Out-Null
    }
    finally {
        Pop-Location
    }

    $wsp = Join-Path $work $solutionName
    if (-not (Test-Path $wsp)) {
        throw 'makecab made no solution'
    }

    $existing = Get-SPSolution $solutionName -ErrorAction SilentlyContinue
    if ($existing -and $existing.Deployed) {
        Write-Host 'Updating the deployed solution: the application pools recycle'
        Update-SPSolution $solutionName -LiteralPath $wsp -GACDeployment
    }
    else {
        if ($existing) {
            Remove-SPSolution $solutionName -Confirm:$false
        }

        Add-SPSolution -LiteralPath $wsp | Out-Null
        Install-SPSolution $solutionName -GACDeployment
    }

    # 2. deployed on every server, before any web.config names the module
    Wait-Solution {
        param($s)
        $s -and $s.Deployed -and -not $s.JobExists -and "$($s.LastOperationResult)" -like '*Succeeded*'
    } 'deployed'
    $servers = @(Get-SPServer | Where-Object {
        $_.Role -ne 'Invalid' -and
        @($_.ServiceInstances | Where-Object { $_.TypeName -like '*Web Application*' -and $_.Status -eq 'Online' }).Count -gt 0
    })
    $solution = Get-SPSolution $solutionName
    Write-Host "Deployed to $($solution.DeployedServers.Count) of $($servers.Count) servers with the web application service"
    if ($solution.DeployedServers.Count -lt $servers.Count) {
        throw 'Not deployed to every server of the web application service: web.config left as it was'
    }
}
finally {
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
}

# 3. the module and its settings in the web.config of each web application
$settings = [ordered] @{
    'OfficeHostShadow.Enabled' = 'true'
    'OfficeHostShadow.Sink' = $Sink
    'OfficeHostShadow.ApiKey' = $ApiKey
    'OfficeHostShadow.SamplePercent' = "$SamplePercent"
    'OfficeHostShadow.MaxInFlightMegabytes' = "$MaxInFlightMegabytes"
    'OfficeHostShadow.MaxExchangeMegabytes' = "$MaxExchangeMegabytes"
    'OfficeHostShadow.Include' = ($Include -join ',')
}
$type = [Microsoft.SharePoint.Administration.SPWebConfigModification+SPWebConfigModificationType]::EnsureChildNode
foreach ($url in $WebApplication) {
    $application = Get-SPWebApplication $url
    Remove-Modifications $application | Out-Null
    $module = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
    $module.Path = 'configuration/system.webServer/modules'
    $module.Name = "add[@name='OfficeHostShadow']"
    $module.Value = "<add name=`"OfficeHostShadow`" type=`"OfficeHost.ShadowModule, $($assembly.FullName)`" />"
    $module.Owner = $owner
    $module.Sequence = 0
    $module.Type = $type
    $application.WebConfigModifications.Add($module)
    foreach ($entry in $settings.GetEnumerator()) {
        $setting = New-Object Microsoft.SharePoint.Administration.SPWebConfigModification
        $setting.Path = 'configuration/appSettings'
        $setting.Name = "add[@key='$($entry.Key)']"
        $setting.Value = "<add key=`"$($entry.Key)`" value=`"$([Security.SecurityElement]::Escape($entry.Value))`" />"
        $setting.Owner = $owner
        $setting.Sequence = 0
        $setting.Type = $type
        $application.WebConfigModifications.Add($setting)
    }

    $application.Update()
}

[Microsoft.SharePoint.Administration.SPWebService]::ContentService.ApplyWebConfigModifications()
Wait-WebConfig

# 4. the event log source of the module, on this server
if (-not [Diagnostics.EventLog]::SourceExists('OfficeHost.Shadow')) {
    New-EventLog -LogName Application -Source 'OfficeHost.Shadow'
}

# what this server has now, as IIS reads it
Import-Module WebAdministration
foreach ($url in $WebApplication) {
    $application = Get-SPWebApplication $url
    $site = $application.IisSettings['Default']
    if (-not (Test-ModuleInWebConfig $application)) {
        throw "The web.config of $url does not have the module yet on this server: see the web.config timer job"
    }

    $modules = Get-WebConfiguration -PSPath "IIS:\Sites\$($site.ServerComment)\_vti_bin" -Filter system.webServer/modules/add | ForEach-Object name
    if ($modules -notcontains 'OfficeHostShadow') {
        throw "IIS does not load the module for $url/_vti_bin"
    }
}

Write-Host "The module of the OfficeHost shadow copies $($WebApplication -join ', ') to $Sink ($SamplePercent% of documents)"
Write-Host 'On every other server, create its event log source: New-EventLog -LogName Application -Source OfficeHost.Shadow'
