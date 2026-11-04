<#
.SYNOPSIS
Sends one cell storage request to SharePoint, WhoAmI and ServerTime (they change nothing), and shows SharePoint's answer
as it came next to the module's copy of it in the recording folder: whether either holds the answer more than once.

.DESCRIPTION
Run on the server of the farm, as a user who can read -Document, while the module records to a folder. Signs in with
your Windows credential. Prints the bytes and SOAP envelopes of the answer as received, then of the newest copy the
module wrote. One envelope in each is right.

.EXAMPLE
.\Test-ShadowWire.ps1 -Site https://sharepoint.example/sites/team -Document 'https://sharepoint.example/sites/team/Shared Documents/x.docx'
#>
param(
    [Parameter(Mandatory)]
    [string] $Site,
    # the URL of a document of the site, as Office names it
    [Parameter(Mandatory)]
    [string] $Document,
    [string] $Recording = 'C:\ShadowRecording'
)

$ErrorActionPreference = 'Stop'

function Count-Envelopes([byte[]] $Bytes) {
    $text = [Text.Encoding]::UTF8.GetString($Bytes)
    ([regex]::Matches($text, '<s:Envelope')).Count
}

$soap = 'http://schemas.microsoft.com/sharepoint/soap/'
$envelope = "<s:Envelope xmlns:s=`"http://schemas.xmlsoap.org/soap/envelope/`"><s:Body><RequestVersion Version=`"2`" MinorVersion=`"2`" xmlns=`"$soap`" /><RequestCollection CorrelationId=`"{$([guid]::NewGuid())}`" xmlns=`"$soap`"><Request Url=`"$([Security.SecurityElement]::Escape($Document))`" RequestToken=`"1`"><SubRequest Type=`"WhoAmI`" SubRequestToken=`"1`" /><SubRequest Type=`"ServerTime`" SubRequestToken=`"2`" /></Request></RequestCollection></s:Body></s:Envelope>"
# MTOM as Word sends it, with a line break before the first boundary
$boundary = "urn:uuid:$([guid]::NewGuid())"
$body = [Text.Encoding]::UTF8.GetBytes("`r`n--$boundary`r`nContent-ID: <http://tempuri.org/xml/0>`r`nContent-Transfer-Encoding: 8bit`r`nContent-Type: application/xop+xml;charset=utf-8;type=`"text/xml; charset=utf-8`"`r`n`r`n$envelope`r`n--$boundary--`r`n")

$before = @(Get-ChildItem $Recording -Filter '*.response.http' -ErrorAction SilentlyContinue).Count
$request = [Net.HttpWebRequest]::Create("$($Site.TrimEnd('/'))/_vti_bin/cellstorage.svc/CellStorageService")
$request.Method = 'POST'
$request.UseDefaultCredentials = $true
$request.ContentType = "multipart/related; type=`"application/xop+xml`"; boundary=`"$boundary`"; start=`"<http://tempuri.org/xml/0>`"; start-info=`"text/xml; charset=utf-8`""
$request.Headers['SOAPAction'] = '"http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest"'
$request.UserAgent = 'Test-ShadowWire'
$stream = $request.GetRequestStream()
$stream.Write($body, 0, $body.Length)
$stream.Close()
try {
    $response = $request.GetResponse()
}
catch [Net.WebException] {
    $response = $_.Exception.Response
    if (-not $response) { throw }
}

$received = New-Object IO.MemoryStream
$response.GetResponseStream().CopyTo($received)
$response.Close()
$wire = $received.ToArray()
Write-Host "SharePoint's answer:  $([int] $response.StatusCode), Content-Length $($response.Headers['Content-Length']), $($wire.Length) bytes, $(Count-Envelopes $wire) envelope(s)"

# the module writes from a thread of its own: give it a few seconds
$copy = $null
for ($i = 0; $i -lt 20 -and -not $copy; $i++) {
    Start-Sleep -Milliseconds 500
    $files = @(Get-ChildItem $Recording -Filter '*.response.http' -ErrorAction SilentlyContinue | Sort-Object Name)
    if ($files.Count -gt $before) {
        $copy = $files[-1]
    }
}

if (-not $copy) {
    Write-Host "The module's copy:    none in $Recording"
    return
}

$bytes = [IO.File]::ReadAllBytes($copy.FullName)
$text = [Text.Encoding]::GetEncoding(28591).GetString($bytes)
$start = $text.IndexOf("`r`n`r`n") + 4
$copied = New-Object byte[] ($bytes.Length - $start)
[Array]::Copy($bytes, $start, $copied, 0, $copied.Length)
Write-Host "The module's copy:    $($copy.Name), $($copied.Length) bytes, $(Count-Envelopes $copied) envelope(s)"
