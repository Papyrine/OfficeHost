<#
.SYNOPSIS
Turns the VM of New-LabVm.ps1 into a one-server SharePoint lab: a domain, SQL Server, SharePoint Server Subscription
Edition, a site, and test users.

.DESCRIPTION
Runs everything in the VM through PowerShell Direct, restarting it where an installation needs it. Each stage checks
whether it is done already, so the script can be run again after a failure, and goes on from there:
1. a domain (lab.test, LAB) with the VM as its domain controller;
2. the test users alice and bob;
3. SQL Server 2022 Developer, the database engine only;
4. the SharePoint prerequisites, and the SharePoint binaries with the trial key;
5. a single server farm, with Central Administration on port 2016;
6. a web application and a team site at http://<Name>.mshome.net, where alice and bob can edit.
LAB\Administrator is the farm account and the account of the application pool: a lab, not a farm to copy. Run
elevated. The passwords are asked for: the Administrator password of New-LabVm.ps1, and one password for the test
users. See docs/sharepoint-lab.md.

.EXAMPLE
src/sharepoint-lab/Install-LabFarm.ps1 -ProductKey <the Enterprise trial key of the download page>
#>
#Requires -RunAsAdministrator
param(
    # the trial key of SharePoint Server Subscription Edition, from its download page
    [Parameter(Mandatory)]
    [string] $ProductKey,
    [string] $Name = 'SPSE',
    [string] $Domain = 'lab.test',
    [string] $NetBios = 'LAB',
    [securestring] $AdministratorPassword = (Read-Host -AsSecureString 'Password of the lab Administrator'),
    [securestring] $UserPassword = (Read-Host -AsSecureString 'Password for the test users alice and bob')
)

$ErrorActionPreference = 'Stop'

# the domain's password policy, which the recovery password of the domain, the farm passphrase and the users' passwords
# have to meet: at least 8 characters, of 3 of uppercase, lowercase, digits and symbols
function Assert-Complex([securestring] $Password, [string] $Of) {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
    $kinds = @('[A-Z]', '[a-z]', '[0-9]', '[^A-Za-z0-9]' | Where-Object { $plain -cmatch $_ }).Count
    $length = $plain.Length
    $plain = $null
    if ($length -lt 8 -or $kinds -lt 3) {
        throw "The password of $Of needs at least 8 characters, of 3 of uppercase, lowercase, digits and symbols"
    }
}

Assert-Complex $AdministratorPassword 'the lab Administrator'
Assert-Complex $UserPassword 'alice and bob'
$adminName = "$NetBios\Administrator"
$local = New-Object PSCredential("$Name\Administrator", $AdministratorPassword)
$domainAdmin = New-Object PSCredential($adminName, $AdministratorPassword)

function Wait-Guest([PSCredential] $Credential, [int] $Minutes = 40) {
    $deadline = (Get-Date).AddMinutes($Minutes)
    while ((Get-Date) -lt $deadline) {
        try {
            Invoke-Command -VMName $Name -Credential $Credential -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop | Out-Null
            return
        }
        catch {
            Start-Sleep -Seconds 15
        }
    }

    throw "$Name did not answer as $($Credential.UserName) within $Minutes minutes"
}

function Invoke-Guest([PSCredential] $Credential, [scriptblock] $Script) {
    Invoke-Command -VMName $Name -Credential $Credential -ScriptBlock $Script -ErrorAction Stop
}

function Restart-Guest([PSCredential] $Credential) {
    Write-Host "Restarting $Name"
    Restart-VM -Name $Name -Force -Wait -For Heartbeat
    Wait-Guest $Credential
}

function Step([string] $Text) {
    Write-Host "== $Text" -ForegroundColor Cyan
}

if ((Get-VM -Name $Name).State -ne 'Running') {
    Start-VM -Name $Name
}

# 1. the domain: the local Administrator answers until the VM is a domain controller
$isController = $false
try {
    $isController = Invoke-Guest $domainAdmin { (Get-CimInstance Win32_ComputerSystem).DomainRole -ge 4 }
}
catch {
}

if (-not $isController) {
    Step "Making $Name the domain controller of $Domain"
    Wait-Guest $local
    Invoke-Guest $local {
        # the answers of the first boot hold the password in plain text
        Remove-Item C:\Windows\Panther\unattend.xml -ErrorAction SilentlyContinue
        Install-WindowsFeature AD-Domain-Services -IncludeManagementTools | Out-Null
        Import-Module ADDSDeployment
        Install-ADDSForest -DomainName $using:Domain -DomainNetbiosName $using:NetBios -SafeModeAdministratorPassword $using:AdministratorPassword -InstallDns -Force -NoRebootOnCompletion -WarningAction SilentlyContinue | Out-Null
    }
    Restart-Guest $domainAdmin
}

# the domain's DNS: with an address of the Default Switch's DHCP, the VM asks the switch's DNS, which does not know the
# domain, so neither the AD cmdlets nor SharePoint find the domain controller. The VM asks itself, and its DNS server
# forwards other names to the switch's.
Step 'Domain DNS'
Invoke-Guest $domainAdmin {
    $adapter = Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1
    $servers = (Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4).ServerAddresses
    if ($servers -notcontains '127.0.0.1') {
        foreach ($server in $servers) {
            Add-DnsServerForwarder -IPAddress $server -ErrorAction SilentlyContinue
        }

        Set-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -ServerAddresses '127.0.0.1'
        Write-Host "DNS: 127.0.0.1, forwarding to $($servers -join ', ')"
    }

    # the promotion of the image of New-LabVm.ps1 left Active Directory Web Services disabled
    if ((Get-Service ADWS).StartType -eq 'Disabled') {
        Set-Service ADWS -StartupType Automatic
        Write-Host 'Active Directory Web Services: was disabled, now automatic'
    }

    foreach ($service in 'DNS', 'NTDS', 'Netlogon', 'ADWS') {
        if ((Get-Service $service).Status -ne 'Running') {
            Start-Service $service
        }
    }

    # the records that locate the domain controller
    Restart-Service Netlogon
    & ipconfig /registerdns | Out-Null
}

# 2. the test users
Step 'Test users'
Invoke-Guest $domainAdmin {
    $domainName = $using:Domain
    $netBiosName = $using:NetBios
    $server = $env:COMPUTERNAME
    # Active Directory Web Services, which the AD cmdlets use, starts minutes after the domain controller does
    $deadline = (Get-Date).AddMinutes(10)
    while ($true) {
        try {
            Get-ADDomain -Server $server -ErrorAction Stop | Out-Null
            break
        }
        catch {
            if ((Get-Date) -gt $deadline) {
                $state = (Get-Service DNS, NTDS, Netlogon, ADWS | ForEach-Object { "$($_.Name) $($_.Status)" }) -join ', '
                throw "Active Directory did not answer: $($_.Exception.Message) ($state)"
            }

            Start-Sleep -Seconds 15
        }
    }

    foreach ($user in 'alice', 'bob') {
        if (-not (Get-ADUser -Server $server -Filter "SamAccountName -eq '$user'")) {
            $display = (Get-Culture).TextInfo.ToTitleCase($user)
            New-ADUser -Server $server -Name $display -DisplayName $display -SamAccountName $user -UserPrincipalName "$user@$domainName" -AccountPassword $using:UserPassword -Enabled $true -PasswordNeverExpires $true
            Write-Host "Created $netBiosName\$user"
        }
    }
}

# 3. SQL Server, from the DVD drive that has setup.exe but no SharePoint prerequisite installer
Step 'SQL Server'
Invoke-Guest $domainAdmin {
    if (Get-Service MSSQLSERVER -ErrorAction SilentlyContinue) {
        Write-Host 'Installed already'
        return
    }

    $drive = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.DriveLetter } |
        Where-Object { (Test-Path "$($_.DriveLetter):\setup.exe") -and -not (Test-Path "$($_.DriveLetter):\prerequisiteinstaller.exe") } |
        Select-Object -First 1
    if (-not $drive) { throw 'No SQL Server media in the DVD drives' }
    $admin = $using:adminName
    $arguments = '/Q', '/ACTION=Install', '/FEATURES=SQLENGINE', '/INSTANCENAME=MSSQLSERVER', "/SQLSYSADMINACCOUNTS=`"$admin`"", '/TCPENABLED=1', '/UPDATEENABLED=False', '/IACCEPTSQLSERVERLICENSETERMS'
    $process = Start-Process "$($drive.DriveLetter):\setup.exe" -ArgumentList $arguments -Wait -PassThru
    if ($process.ExitCode -notin 0, 3010) { throw "SQL Server setup failed: $($process.ExitCode)" }
}

# 4. the SharePoint prerequisites (restarting as they ask), then the binaries
$binaries = { Test-Path "$env:CommonProgramFiles\microsoft shared\Web Server Extensions\16\BIN\psconfig.exe" }
if (-not (Invoke-Guest $domainAdmin $binaries)) {
    Step 'SharePoint prerequisites'
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        $code = Invoke-Guest $domainAdmin {
            $drive = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.DriveLetter } |
                Where-Object { Test-Path "$($_.DriveLetter):\prerequisiteinstaller.exe" } |
                Select-Object -First 1
            if (-not $drive) { throw 'No SharePoint media in the DVD drives' }
            (Start-Process "$($drive.DriveLetter):\prerequisiteinstaller.exe" -ArgumentList '/unattended' -Wait -PassThru).ExitCode
        }
        Write-Host "Prerequisite installer: $code"
        if ($code -eq 0) {
            break
        }

        # 3010: restart; 1001: a restart is pending, then run it again
        if ($code -notin 1001, 3010) { throw "The prerequisite installer failed: $code" }
        Restart-Guest $domainAdmin
    }

    Step 'SharePoint binaries'
    $config = (Get-Content (Join-Path $PSScriptRoot 'sharepoint-config.xml') -Raw).Replace('{ProductKey}', $ProductKey)
    Invoke-Guest $domainAdmin {
        New-Item -ItemType Directory -Force C:\Lab | Out-Null
        Set-Content C:\Lab\sharepoint-config.xml $using:config
        $drive = Get-Volume | Where-Object { $_.DriveType -eq 'CD-ROM' -and $_.DriveLetter } |
            Where-Object { Test-Path "$($_.DriveLetter):\prerequisiteinstaller.exe" } |
            Select-Object -First 1
        # SharePoint SE takes the license on the command line: AcceptEula of the config is not enough (30203)
        $process = Start-Process "$($drive.DriveLetter):\setup.exe" -ArgumentList '/config', 'C:\Lab\sharepoint-config.xml', '/IAcceptTheLicenseTerms' -Wait -PassThru
        Remove-Item C:\Lab\sharepoint-config.xml
        if ($process.ExitCode -ne 0) { throw "SharePoint setup failed: $($process.ExitCode); its log is in %temp% of the VM" }
    }
    Restart-Guest $domainAdmin
}

# 5. the farm, and 6. the site
Step 'Farm, web application and site'
$url = "http://$($Name.ToLowerInvariant()).mshome.net"
Invoke-Guest $domainAdmin {
    $admin = $using:adminName
    $netBiosName = $using:NetBios
    $credential = $using:domainAdmin
    try { $Host.Runspace.ThreadOptions = 'ReuseThread' } catch { }
    Import-Module SharePointServer -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (-not (Get-Command New-SPConfigurationDatabase -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
    }

    $farm = $null
    try { $farm = Get-SPFarm -ErrorAction Stop } catch { }
    if (-not $farm) {
        # SQL Server takes its connections a while after a restart of the VM, later than PowerShell Direct does
        $deadline = (Get-Date).AddMinutes(10)
        while ($true) {
            $connection = New-Object System.Data.SqlClient.SqlConnection "Data Source=$env:COMPUTERNAME;Initial Catalog=master;Integrated Security=True;Pooling=False"
            try {
                $connection.Open()
                break
            }
            catch {
                if ((Get-Date) -gt $deadline) { throw "SQL Server did not take connections: $($_.Exception.InnerException.Message)" }
                Start-Sleep -Seconds 10
            }
            finally {
                $connection.Dispose()
            }
        }

        Write-Host 'Creating the farm'
        New-SPConfigurationDatabase -DatabaseName SP_Config -DatabaseServer $env:COMPUTERNAME -AdministrationContentDatabaseName SP_Admin -Passphrase $using:AdministratorPassword -FarmCredentials $credential -LocalServerRole SingleServerFarm
    }

    # the rest of the configuration wizard, each step safe to run again
    Write-Host 'Configuring the farm'
    Install-SPHelpCollection -All
    Initialize-SPResourceSecurity
    # the service instances: on a slow VM, this times out with some of them left to do, which a run again does
    try {
        Install-SPService -ErrorAction Stop
    }
    catch {
        Write-Warning "Install-SPService: $($_.Exception.Message) Run the script again to finish it."
    }

    Install-SPFeature -AllExistingFeatures -Force | Out-Null
    if (-not (Get-SPWebApplication -IncludeCentralAdministration | Where-Object IsAdministrationWebApplication)) {
        Write-Host 'Creating Central Administration'
        New-SPCentralAdministration -Port 2016 -WindowsAuthProvider NTLM
    }

    Install-SPApplicationContent

    $url = $using:url
    if (-not (Get-SPWebApplication $url -ErrorAction SilentlyContinue)) {
        Write-Host "Creating the web application $url"
        $account = Get-SPManagedAccount $admin -ErrorAction SilentlyContinue
        if (-not $account) {
            $account = New-SPManagedAccount $credential
        }

        New-SPWebApplication -Name Lab -URL $url -Port 80 -HostHeader ([uri] $url).Host -ApplicationPool Lab -ApplicationPoolAccount $account -AuthenticationProvider (New-SPAuthenticationProvider) -DatabaseName WSS_Content_Lab | Out-Null
    }

    if (-not (Get-SPSite $url -ErrorAction SilentlyContinue)) {
        Write-Host "Creating the site $url"
        New-SPSite $url -OwnerAlias $admin -Name Lab -Template STS#3 | Out-Null
    }

    foreach ($user in 'alice', 'bob') {
        New-SPUser -UserAlias "$netBiosName\$user" -Web $url -PermissionLevel Contribute -ErrorAction SilentlyContinue | Out-Null
    }
}

Write-Host "The lab is at $url (Documents: $url/Shared%20Documents). Sign in as $NetBios\alice or $NetBios\bob; Central Administration is http://$($Name.ToLowerInvariant()).mshome.net:2016." -ForegroundColor Green
