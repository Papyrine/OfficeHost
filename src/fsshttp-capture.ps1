<#
.SYNOPSIS
Captures the HTTP traffic of desktop Office with a document server, decrypted, without a proxy or a root certificate.

.DESCRIPTION
Starts an ETW trace of the Windows HTTP stacks Office uses (WinINet, and WinHTTP through WebIO), opens the document in
Word, Excel or PowerPoint (ms-word:, ms-excel: or ms-powerpoint:, by its extension), or leaves it to be opened from the
site with -HostName alone, and stops the trace when Enter is pressed (or after -Seconds). Then writes the requests and
responses to the document host as text: the headers (WebIO 100/101, WinINet 210/211) and the bytes sent and received
(WebIO 111/122), which ETW logs before encryption. Authorization, cookie and ticket values are redacted.

Starting the trace needs membership of Performance Log Users (or an elevated prompt). The trace (.etl) holds
credentials: keep it private, and share the text output only. See docs/fsshttp-spike.md.

.EXAMPLE
pwsh src/fsshttp-capture.ps1 -Url 'https://sharepoint.example/Shared%20Documents/test.docx'

.EXAMPLE
pwsh src/fsshttp-capture.ps1 -HostName sharepoint.example -Seconds 300

.EXAMPLE
pwsh src/fsshttp-capture.ps1 -Etl capture.etl -HostName sharepoint.example
#>
param(
    # the document to open in Office
    [string] $Url,
    # an existing trace to decode, instead of capturing
    [string] $Etl,
    # the host whose traffic to keep, by default the host of -Url
    [string] $HostName,
    # how long to capture, 0 to wait for Enter
    [int] $Seconds = 0,
    [string] $Out = 'fsshttp-capture.txt'
)

$ErrorActionPreference = 'Stop'
$session = 'FssHttpCapture'

if (-not $Etl) {
    if (-not $Url -and -not $HostName) { throw 'Pass -Url (the document to open), -HostName (to open it yourself) or -Etl (a trace to decode)' }
    $Etl = Join-Path (Get-Location) "fsshttp-capture-$(Get-Date -Format 'yyyyMMdd-HHmmss').etl"
    # not WinINet-Capture, whose bodies need an elevated prompt: WebIO logs those of WinHTTP, which Word uses for documents
    logman stop $session -ets 2>$null | Out-Null
    logman start $session -ets -o $Etl -p Microsoft-Windows-WebIO 0xffffffffffffffff 0xff -bs 1024 -nb 64 512 -max 1024 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Could not start the trace: run from a member of Performance Log Users or elevated" }
    foreach ($provider in 'Microsoft-Windows-WinINet', 'Microsoft-Windows-WinHttp') {
        logman update $session -ets -p $provider 0xffffffffffffffff 0xff | Out-Null
    }
    try {
        if ($Url) {
            $extension = [IO.Path]::GetExtension(([uri] $Url).AbsolutePath).ToLowerInvariant()
            $scheme = 'ms-word'
            if ($extension -in '.xlsx', '.xlsm', '.xlsb', '.xls') { $scheme = 'ms-excel' }
            if ($extension -in '.pptx', '.pptm', '.ppt') { $scheme = 'ms-powerpoint' }
            Start-Process "${scheme}:ofe|u|$Url"
        }
        if ($Seconds -gt 0) {
            Write-Host "Capturing for $Seconds seconds"
            Start-Sleep $Seconds
        }
        else {
            Read-Host 'Capturing: use Office, then press Enter to stop'
        }
    }
    finally {
        logman stop $session -ets | Out-Null
    }
}

if (-not $HostName -and $Url) { $HostName = ([uri] $Url).Host }

function Redact([string] $text) {
    # keeps the scheme or the cookie name only
    $text -replace '(?im)^((?:Authorization|Proxy-Authorization|Cookie|Set-Cookie|AuthMsaDeviceTicket|X-IDCRL_AUTH_PARAMS_V1)\s*:\s*[\w-]+)[^\r\n]{8,}', '$1 <redacted>'
}

# text when the bytes are mostly printable, a hex dump otherwise
function Format-Bytes([byte[]] $bytes) {
    $printable = ($bytes | Where-Object { ($_ -ge 0x20 -and $_ -lt 0x7f) -or $_ -in 9, 10, 13 }).Count
    if ($bytes.Length -gt 0 -and $printable / $bytes.Length -ge 0.9) {
        return Redact ([Text.Encoding]::UTF8.GetString($bytes))
    }
    $lines = for ($offset = 0; $offset -lt [Math]::Min($bytes.Length, 4096); $offset += 16) {
        $line = $bytes[$offset..([Math]::Min($offset + 15, $bytes.Length - 1))]
        $hex = ($line | ForEach-Object { $_.ToString('x2') }) -join ' '
        $ascii = -join ($line | ForEach-Object { if ($_ -ge 0x20 -and $_ -lt 0x7f) { [char] $_ } else { '.' } })
        '{0:x8}  {1,-47}  {2}' -f $offset, $hex, $ascii
    }
    if ($bytes.Length -gt 4096) { $lines += "... $($bytes.Length) bytes in all" }
    $lines -join "`n"
}

Write-Host "Decoding $Etl"
$events = Get-WinEvent -Path $Etl -Oldest -ErrorAction SilentlyContinue |
    Where-Object { $_.ProviderName -in 'Microsoft-Windows-WebIO', 'Microsoft-Windows-WinINet' }

function Keep([string] $host_) {
    -not $HostName -or $host_ -eq $HostName
}

# WebIO logs per request (17: its URI, 104: its connection) and per connection (122: the bytes received). Their
# addresses are reused, so the maps are kept in time order, as the events come.
$requestHost = @{}
$connectionHost = @{}
$winInetHost = @{}
$output = foreach ($e in $events) {
    $prefix = [regex]::Match($e.Message, '^(0x[0-9a-fA-F]+):').Groups[1].Value
    $key = "$($e.ProcessId)/$prefix"
    $time = $e.TimeCreated.ToString('HH:mm:ss.fff')
    $bytes = $e.Properties | Where-Object { $_.Value -is [byte[]] } | Select-Object -First 1
    switch ("$($e.ProviderName)/$($e.Id)") {
        'Microsoft-Windows-WebIO/17' {
            $requestHost[$key] = ([uri] [regex]::Match($e.Message, '\(URI (\S+)\)').Groups[1].Value).Host
        }
        'Microsoft-Windows-WebIO/104' {
            $connection = [regex]::Match($e.Message, '\(Connection: (0x[0-9a-fA-F]+)\)').Groups[1].Value
            $connectionHost["$($e.ProcessId)/$connection"] = $requestHost[$key]
        }
        'Microsoft-Windows-WebIO/100' {
            if (Keep $requestHost[$key]) { "==== $time pid $($e.ProcessId) request`n" + (Redact ($e.Message -replace '^.*?Sending Headers: ', '')) }
        }
        'Microsoft-Windows-WebIO/101' {
            if (Keep $requestHost[$key]) { "==== $time pid $($e.ProcessId) response`n" + (Redact ($e.Message -replace '^.*?Received Headers: ', '')) }
        }
        'Microsoft-Windows-WebIO/111' {
            if ($bytes -and (Keep $requestHost[$key])) { "---- $time pid $($e.ProcessId) sent bytes`n" + (Format-Bytes $bytes.Value) }
        }
        'Microsoft-Windows-WebIO/122' {
            if ($bytes -and (Keep $connectionHost[$key])) { "---- $time pid $($e.ProcessId) received bytes`n" + (Format-Bytes $bytes.Value) }
        }
        { $_ -in 'Microsoft-Windows-WinINet/210', 'Microsoft-Windows-WinINet/211' } {
            $handle = "$($e.ProcessId)/$([regex]::Match($e.Message, 'RequestHandle=(0x[0-9a-fA-F]+)').Groups[1].Value)"
            if ($e.Id -eq 210) {
                $winInetHost[$handle] = [regex]::Match($e.Message, '(?m)^Host: (\S+)').Groups[1].Value
            }
            if (Keep $winInetHost[$handle]) {
                "==== $time pid $($e.ProcessId) WinINet`n" + (Redact ($e.Message -replace '^HTTP (Request|Response) Headers RequestHandle=\S+ ', ''))
            }
        }
    }
}

$output | Set-Content -Path $Out -Encoding utf8
Write-Host "Wrote $Out ($(@($output).Count) entries)"
