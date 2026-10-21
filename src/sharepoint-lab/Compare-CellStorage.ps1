<#
.SYNOPSIS
Sends the same cell storage sub-requests to SharePoint (the lab) and to OfficeHost (the sample), and shows where their
answers differ.

.DESCRIPTION
Each scenario is a few requests of XML sub-requests (locks, coauthoring, dependencies, errors) on one document, with
client and lock ids of its own, then clean-up requests that release what it took. The answers are compared by their
ErrorCode and their main attributes (LockType, CoauthStatus, ExclusiveLockReturnReason, AmIAlone, ErrorMessage). See
docs/sharepoint-lab.md.

SharePoint is signed in to with -Credential, or else with the Windows credential saved for its host (cmdkey). The
sample runs with its desktop launch profile, with -File in its files, and is signed in to as user alice of its tokens.

.EXAMPLE
pwsh src/sharepoint-lab/Compare-CellStorage.ps1 -Credential (Get-Credential LAB\alice)
#>
param(
    [pscredential] $Credential,
    [string] $SharePoint = 'http://spse.mshome.net',
    [string] $Library = 'Shared%20Documents',
    [string] $Sample = 'https://localhost:5001',
    [string] $File = 'Lab.docx',
    # only the scenarios whose name matches
    [string] $Filter = '.'
)

$ErrorActionPreference = 'Stop'
$soap = 'http://schemas.microsoft.com/sharepoint/soap/'
$schemaLockId = '{29358EC1-E813-4793-8E70-ED0344E7B73C}'

# the two targets: where to post, the Url of the document, and how to sign in
$sampleSite = $null
$page = (Invoke-WebRequest "$Sample/" -UseBasicParsing).Content
foreach ($id in [regex]::Matches($page, 'href="/desktop/([^"]+)"') | ForEach-Object { $_.Groups[1].Value }) {
    $html = (Invoke-WebRequest "$Sample/desktop/$id`?user=alice" -UseBasicParsing).Content
    $link = [System.Net.WebUtility]::HtmlDecode([regex]::Match($html, 'id="open" href="([^"]+)"').Groups[1].Value)
    if ($link -like "*/$File") {
        $sampleSite = [regex]::Match($link, 'https?://.+/fsshttp/[^/]+').Value
        break
    }
}
if (-not $sampleSite) { throw "No $File in the sample's files" }

$targets = @(
    @{ Name = 'SharePoint'; Endpoint = "$SharePoint/_vti_bin/cellstorage.svc/CellStorageService"; Folder = "$SharePoint/$Library" },
    @{ Name = 'OfficeHost'; Endpoint = "$sampleSite/_vti_bin/cellstorage.svc/CellStorageService"; Folder = $sampleSite }
)

# the TransitionID of the last join of each target, for {TransitionID} in the sub-requests after it
$transitions = @{}

# $name: the file the request is for, -File by default, another after a rename
function Send($target, [string[]] $subRequests, [int] $version = 2, [string] $name = $File) {
    $subRequests = $subRequests | ForEach-Object { $_.Replace('{TransitionID}', "$($transitions[$target.Name])") }
    $envelope ="<s:Envelope xmlns:s=`"http://schemas.xmlsoap.org/soap/envelope/`"><s:Body><RequestVersion Version=`"$version`" MinorVersion=`"2`" xmlns=`"$soap`" /><RequestCollection CorrelationId=`"{$([guid]::NewGuid())}`" xmlns=`"$soap`"><Request Url=`"$($target.Folder)/$name`" RequestToken=`"1`">$($subRequests -join '')</Request></RequestCollection></s:Body></s:Envelope>"
    # MTOM as Word sends it, with a line break before the first boundary: without it, SharePoint redirects to its error
    # page ("Object reference not set to an instance of an object")
    $boundary = "urn:uuid:$([guid]::NewGuid())"
    $body = "`r`n--$boundary`r`nContent-ID: <http://tempuri.org/xml/0>`r`nContent-Transfer-Encoding: 8bit`r`nContent-Type: application/xop+xml;charset=utf-8;type=`"text/xml; charset=utf-8`"`r`n`r`n$envelope`r`n--$boundary--`r`n"
    $arguments = @{
        Uri = $target.Endpoint
        Method = 'Post'
        Body = [Text.Encoding]::UTF8.GetBytes($body)
        ContentType = "multipart/related; type=`"application/xop+xml`"; boundary=`"$boundary`"; start=`"<http://tempuri.org/xml/0>`"; start-info=`"text/xml; charset=utf-8`""
        Headers = @{ SOAPAction = '"http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest"' }
        UseBasicParsing = $true
        SkipHttpErrorCheck = $true
    }
    if ($target.Name -eq 'SharePoint') {
        $arguments.AllowUnencryptedAuthentication = $true
        if ($Credential) { $arguments.Credential = $Credential } else { $arguments.UseDefaultCredentials = $true }
    }
    $response = Invoke-WebRequest @arguments
    $text = $response.Content
    if ($text -is [byte[]]) {
        $text = [Text.Encoding]::UTF8.GetString($text)
    }

    if ($response.StatusCode -ne 200 -or $text -notmatch 'ResponseVersion') {
        $error = [regex]::Match($text, 'ErrorText=([^"&]+)').Groups[1].Value
        return @("HTTP $($response.StatusCode) $([Uri]::UnescapeDataString($error.Replace('+', ' ')))".Trim())
    }

    $transition = [regex]::Match($text, 'TransitionID="([^"]+)"')
    if ($transition.Success) {
        $transitions[$target.Name] = $transition.Groups[1].Value
    }

    $results = foreach ($match in [regex]::Matches($text, '<SubResponse ([^>]*?)/?>(?:<SubResponseData([^>]*?)/?>)?')) {
        $outer = $match.Groups[1].Value
        $data = $match.Groups[2].Value
        $code = [regex]::Match($outer, 'ErrorCode="(\w+)"').Groups[1].Value
        $token = [regex]::Match($outer, 'SubRequestToken="(\w+)"').Groups[1].Value
        $message = [regex]::Match($outer, 'ErrorMessage="([^"]*)"').Groups[1].Value
        $attributes = ([regex]::Matches($data, '(LockType|CoauthStatus|ExclusiveLockReturnReason|AmIAlone)="([^"]*)"') | ForEach-Object { "$($_.Groups[1].Value)=$($_.Groups[2].Value)" }) -join ' '
        if ($message) { $attributes = "$attributes ErrorMessage=$message".Trim() }
        "#$token $code $attributes".Trim()
    }
    if (-not $results) {
        # a Response with an error and no SubResponse, as for a bad version
        $results = @(([regex]::Matches($text, 'ErrorCode="(\w+)"') | ForEach-Object { $_.Groups[1].Value }) -join ',')
    }
    @($results)
}

function SubRequest([string] $type, [int] $token, [string] $data, [string] $dependsOn, [string] $dependencyType) {
    $dependency = ''
    if ($dependsOn) { $dependency = " DependsOn=`"$dependsOn`" DependencyType=`"$dependencyType`"" }
    if (-not $data) { return "<SubRequest Type=`"$type`" SubRequestToken=`"$token`"$dependency />" }
    "<SubRequest Type=`"$type`" SubRequestToken=`"$token`"$dependency><SubRequestData $data /></SubRequest>"
}

function Coauth([string] $type, [string] $client, [string] $schema = $schemaLockId, [string] $extra = '') {
    SubRequest 'Coauth' 1 "CoauthRequestType=`"$type`" SchemaLockID=`"$schema`" ClientID=`"$client`" Timeout=`"3600`" $extra"
}

function Schema([string] $type, [string] $client) {
    SubRequest 'SchemaLock' 1 "SchemaLockRequestType=`"$type`" SchemaLockID=`"$schemaLockId`" ClientID=`"$client`" Timeout=`"3600`""
}

function Exclusive([string] $type, [string] $lockId, [string] $extra = '') {
    SubRequest 'ExclusiveLock' 1 "ExclusiveLockRequestType=`"$type`" ExclusiveLockID=`"$lockId`" Timeout=`"600`" $extra"
}

function Scenario([string] $name, [scriptblock] $steps, [scriptblock] $cleanup = { @() }) {
    if ($name -notmatch $Filter) { return }
    $a = "{$([guid]::NewGuid().ToString().ToUpperInvariant())}"
    $b = "{$([guid]::NewGuid().ToString().ToUpperInvariant())}"
    $e = "{$([guid]::NewGuid().ToString().ToUpperInvariant())}"
    $answers = @{}
    foreach ($target in $targets) {
        # a step is a sub-request, an array of them, or @{ SubRequests; Version; File }
        $answers[$target.Name] = @(foreach ($step in & $steps $a $b $e) {
            $version = 2
            $fileName = $File
            if ($step -is [hashtable]) {
                if ($step.Version) { $version = $step.Version }
                if ($step.File) { $fileName = $step.File }
                $step = $step.SubRequests
            }
            (Send $target @($step) $version $fileName) -join '; '
        })
        foreach ($step in & $cleanup $a $b $e) {
            $fileName = $File
            if ($step -is [hashtable]) {
                if ($step.File) { $fileName = $step.File }
                $step = $step.SubRequests
            }
            Send $target @($step) 2 $fileName | Out-Null
        }
    }

    Write-Host "== $name" -ForegroundColor Cyan
    for ($i = 0; $i -lt $answers['SharePoint'].Count; $i++) {
        $sharePoint = $answers['SharePoint'][$i]
        $officeHost = $answers['OfficeHost'][$i]
        if ($sharePoint -ceq $officeHost) {
            Write-Host "  $($i + 1). same: $sharePoint"
        }
        else {
            Write-Host "  $($i + 1). SharePoint: $sharePoint" -ForegroundColor Yellow
            Write-Host "     OfficeHost: $officeHost" -ForegroundColor Yellow
        }
    }
}

# coauthoring sessions
Scenario 'Exit twice' { param($a) (Coauth JoinCoauthoring $a), (Coauth ExitCoauthoring $a), (Coauth ExitCoauthoring $a) }
Scenario 'Exit of a client not in the session' { param($a, $b) (Coauth JoinCoauthoring $a), (Coauth ExitCoauthoring $b) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Exit with another schema lock id' { param($a) (Coauth JoinCoauthoring $a), (Coauth ExitCoauthoring $a '{11111111-2222-3333-4444-555555555555}') } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Status without a session' { param($a) (Coauth GetCoauthoringStatus $a) }
Scenario 'Refresh without a session' { param($a) (Coauth RefreshCoauthoring $a) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Mark transition complete without a session' { param($a) (Coauth MarkTransitionComplete $a) }
Scenario 'Join without a schema lock id' { param($a) (SubRequest 'Coauth' 1 "CoauthRequestType=`"JoinCoauthoring`" ClientID=`"$a`" Timeout=`"3600`"") } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Coauth without a request type' { param($a) (SubRequest 'Coauth' 1 "SchemaLockID=`"$schemaLockId`" ClientID=`"$a`"") }
Scenario 'Two clients: status, AmIAlone, LockStatus' { param($a, $b) (Coauth JoinCoauthoring $a), (Coauth JoinCoauthoring $b), (Coauth GetCoauthoringStatus $a), (SubRequest 'AmIAlone' 1 $null), (SubRequest 'LockStatus' 1 $null) } { param($a, $b) (Coauth ExitCoauthoring $a), (Coauth ExitCoauthoring $b) }
Scenario 'Editor metadata without a join' { param($a) (SubRequest 'EditorsTable' 1 "EditorsTableRequestType=`"UpdateEditorMetadata`" ClientID=`"$a`" Key=`"k`" Value=`"AA==`"") }

# schema and exclusive locks
Scenario 'Schema lock release without a lock' { param($a) (Schema ReleaseLock $a) }
Scenario 'Exclusive release without a lock' { param($a, $b, $e) (Exclusive ReleaseLock $e) }
Scenario 'Exclusive convert without a lock' { param($a, $b, $e) (Exclusive ConvertToSchema $e "SchemaLockID=`"$schemaLockId`" ClientID=`"$a`"") }
Scenario 'Exclusive without a request type' { param($a, $b, $e) (SubRequest 'ExclusiveLock' 1 "ExclusiveLockID=`"$e`"") }
Scenario 'Exclusive, then a join' { param($a, $b, $e) (Exclusive GetLock $e), (Coauth JoinCoauthoring $a), (SubRequest 'LockStatus' 1 $null) } { param($a, $b, $e) (Exclusive ReleaseLock $e) }
Scenario 'Join, then an exclusive lock' { param($a, $b, $e) (Coauth JoinCoauthoring $a), (Exclusive GetLock $e) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Lock status and AmIAlone without a lock' { (SubRequest 'LockStatus' 1 $null), (SubRequest 'AmIAlone' 1 $null) }

# dependencies and requests
Scenario 'OnExecute on a sub-request that did not run' { ,@((SubRequest 'WhoAmI' 1 $null), (SubRequest 'ServerTime' 2 $null '1' 'OnFail'), (SubRequest 'ServerTime' 3 $null '2' 'OnExecute')) }
Scenario 'OnExecute on a sub-request skipped OnNotSupported' { ,@((SubRequest 'WhoAmI' 1 $null), (SubRequest 'ServerTime' 2 $null '1' 'OnNotSupported'), (SubRequest 'ServerTime' 3 $null '2' 'OnExecute')) }
Scenario 'OnExecute on a sub-request skipped OnSuccess' { param($a) ,@((SubRequest 'EditorsTable' 1 "EditorsTableRequestType=`"UpdateEditorMetadata`" ClientID=`"$a`" Key=`"k`" Value=`"AA==`""), (SubRequest 'ServerTime' 2 $null '1' 'OnSuccess'), (SubRequest 'ServerTime' 3 $null '2' 'OnExecute'), (SubRequest 'ServerTime' 4 $null '1' 'OnExecute')) }
Scenario 'OnExecute on a sub-request that failed' { param($a, $b, $e) ,@((Exclusive ReleaseLock $e), (SubRequest 'ServerTime' 2 $null '1' 'OnExecute')) }
Scenario 'Exit with another schema lock id, then status' { param($a) (Coauth JoinCoauthoring $a), (Coauth ExitCoauthoring $a '{11111111-2222-3333-4444-555555555555}'), (Coauth GetCoauthoringStatus $a), (SubRequest 'LockStatus' 1 $null) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Dependency on an unknown token' { ,@((SubRequest 'ServerTime' 1 $null '9' 'OnSuccess')) }
Scenario 'An unknown sub-request type' { ,@((SubRequest 'Bogus' 1 $null)) }
Scenario 'Request version 1' { @{ Version = 1; SubRequests = @((SubRequest 'WhoAmI' 1 $null)) } }
Scenario 'GetVersions' { ,@((SubRequest 'GetVersions' 1 $null)) }

# AmIAlone with the TransitionID of the session, and with another
Scenario 'AmIAlone with the TransitionID' { param($a, $b) (Coauth JoinCoauthoring $a), (SubRequest 'AmIAlone' 1 'TransitionID="{TransitionID}"'), (Coauth JoinCoauthoring $b), (SubRequest 'AmIAlone' 1 'TransitionID="{TransitionID}"'), (SubRequest 'AmIAlone' 1 'TransitionID="{00000000-0000-0000-0000-000000000001}"') } { param($a, $b) (Coauth ExitCoauthoring $a), (Coauth ExitCoauthoring $b) }

# coauthoring timeouts, the spec's range being 60 to 120000
function Join([string] $client, [int] $timeout) {
    SubRequest 'Coauth' 1 "CoauthRequestType=`"JoinCoauthoring`" SchemaLockID=`"$schemaLockId`" ClientID=`"$client`" Timeout=`"$timeout`""
}
Scenario 'Join with a timeout of 30' { param($a) (Join $a 30) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Join with a timeout of 60' { param($a) (Join $a 60) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Join with a timeout of 3599' { param($a) (Join $a 3599) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Join with a timeout of 120000' { param($a) (Join $a 120000) } { param($a) (Coauth ExitCoauthoring $a) }
Scenario 'Join with a timeout of 120001' { param($a) (Join $a 120001) } { param($a) (Coauth ExitCoauthoring $a) }

# the number of coauthors
$coauthors = 1..12 | ForEach-Object { "{$([guid]::NewGuid().ToString().ToUpperInvariant())}" }
Scenario 'Twelve coauthors' { $coauthors | ForEach-Object { Coauth JoinCoauthoring $_ } } { $coauthors | ForEach-Object { Coauth ExitCoauthoring $_ } }

# Cell queries: Excel's query of the content, with its attributes
$contentQuery = 'DgALAJzPKfM5lAabBgIAAO4CAACqAiAABvWzkcnfuk2ORYXED+IRNloEFgANbXNleGNsB3dpbnoCCABm9qA5dwFqBAIAARYCBgADBQCKAgQAAEjaAgYAAwAAygIIAAgAgAOEAEELAawCAFUDAQ=='
function CellQuery([string] $extra = '') {
    "<SubRequest Type=`"Cell`" SubRequestToken=`"1`"><SubRequestData GetFileProps=`"true`" BinaryDataSize=`"109`" $extra>$contentQuery</SubRequestData></SubRequest>"
}
Scenario 'Cell query' { (CellQuery) }
Scenario 'Cell query with another Etag' { (CellQuery 'Etag="&quot;{00000000-0000-0000-0000-000000000001},1&quot;"') }
Scenario 'Cell query expecting no file' { (CellQuery 'ExpectNoFileExists="true"') }

# Properties and Versioning
Scenario 'Properties enumerate' { (SubRequest 'Properties' 1 'Properties="PropertyEnumerate"') }
Scenario 'Properties get' { "<SubRequest Type=`"Properties`" SubRequestToken=`"1`"><SubRequestData Properties=`"PropertyGet`"><PropertyIds><PropertyId id=`"RealtimeTypingEndpointUrl`" /></PropertyIds></SubRequestData></SubRequest>" }
Scenario 'Versioning version list' { (SubRequest 'Versioning' 1 'VersioningRequestType="GetVersionList"') }
Scenario 'Versioning restore of a version that does not exist' { (SubRequest 'Versioning' 1 'VersioningRequestType="RestoreVersion" Version="99.0"') }

# renames, each renamed back: the steps after a rename name the file by its new name
function Rename([string] $newName, [string] $extra = '') {
    SubRequest 'FileOperation' 1 "FileOperation=`"Rename`" NewFileName=`"$newName`" $extra"
}
$renamed = 'LabRenamed.docx'
Scenario 'Rename without a lock' { (Rename $renamed), @{ File = $renamed; SubRequests = @(Rename $File) } } { @{ File = $renamed; SubRequests = @(Rename $File) } }
Scenario 'Rename while in a coauthoring session' { param($a) (Coauth JoinCoauthoring $a), (Rename $renamed), @{ File = $renamed; SubRequests = @(Rename $File) } } { param($a) @{ File = $renamed; SubRequests = @(Rename $File) }, (Coauth ExitCoauthoring $a) }
Scenario 'Rename under an exclusive lock, with its id' { param($a, $b, $e) (Exclusive GetLock $e), (Rename $renamed "ExclusiveLockID=`"$e`""), @{ File = $renamed; SubRequests = @(Rename $File "ExclusiveLockID=`"$e`"") } } { param($a, $b, $e) @{ File = $renamed; SubRequests = @(Rename $File "ExclusiveLockID=`"$e`"") }, (Exclusive ReleaseLock $e) }
Scenario 'Rename under an exclusive lock, without its id' { param($a, $b, $e) (Exclusive GetLock $e), (Rename $renamed) } { param($a, $b, $e) @{ File = $renamed; SubRequests = @(Rename $File "ExclusiveLockID=`"$e`"") }, (Exclusive ReleaseLock $e) }
Scenario 'Rename to another extension' { (Rename 'LabRenamed.txt'), @{ File = 'LabRenamed.txt'; SubRequests = @(Rename $File) } } { @{ File = 'LabRenamed.txt'; SubRequests = @(Rename $File) } }
Scenario 'Rename to an illegal name' { (Rename 'Lab?.docx') }
Scenario 'Rename onto another file' { (Rename 'Lab.pptx') }
Scenario 'Rename without a name' { (SubRequest 'FileOperation' 1 "FileOperation=`"Rename`"") }
Scenario 'File operation without a type' { (SubRequest 'FileOperation' 1 "NewFileName=`"$renamed`"") } { @{ File = $renamed; SubRequests = @(Rename $File) } }
