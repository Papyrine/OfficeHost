<#
.SYNOPSIS
Sends a recording of the module (docs/replay.md) to a shadow, as the module would have sent it live: to try a shadow
before pointing SharePoint at it, or to send it what it missed.

.DESCRIPTION
Posts the exchanges of the folder (0001.request.http and 0001.response.http, ...) in their order, in batches of the
protocol of docs/shadow.md, as one instance of the module, the last batch final. Each exchange is replayed at the time
it was recorded (its X-Replay-Time). Needs PowerShell 7.

.EXAMPLE
.\Send-Recording.ps1 -Folder C:\ShadowRecordings\2026-11-05-word -ApiKey <key> -SkipCertificateCheck
#>
param(
    [Parameter(Mandatory)]
    [string] $Folder,
    [string] $Url = 'https://localhost:5003/officehost-shadow/exchanges',
    [Parameter(Mandatory)]
    [string] $ApiKey,
    [string] $Server = [Environment]::MachineName,
    [int] $BatchMegabytes = 64,
    # for a shadow on localhost with the development certificate
    [switch] $SkipCertificateCheck
)

$ErrorActionPreference = 'Stop'

function Get-Started([byte[]] $Request) {
    # the head of the request, up to its blank line
    $text = [Text.Encoding]::ASCII.GetString($Request, 0, [Math]::Min($Request.Length, 16384))
    $head = $text.Substring(0, [Math]::Max(0, $text.IndexOf("`r`n`r`n")))
    foreach ($line in $head -split "`r`n") {
        if ($line -match '^X-Replay-Time:\s*(.+)$') {
            return [DateTimeOffset]::Parse($Matches[1].Trim(), [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    throw "No X-Replay-Time in the request"
}

$exchanges = @(Get-ChildItem -LiteralPath $Folder -Filter '*.request.http' | Sort-Object Name | ForEach-Object {
    $number = [long] ($_.Name -replace '\..*$', '')
    $response = Join-Path $Folder "$($_.Name -replace '\.request\.http$', '').response.http"
    if (-not (Test-Path -LiteralPath $response)) {
        Write-Warning "No response for $($_.Name): not sent"
        return
    }
    $request = [IO.File]::ReadAllBytes($_.FullName)
    [pscustomobject] @{
        Number = $number
        Started = Get-Started $request
        Request = $request
        Response = [IO.File]::ReadAllBytes($response)
    }
})
if ($exchanges.Count -eq 0) {
    throw "No exchanges in $Folder"
}

# batches of at most $BatchMegabytes, an exchange at least
$batches = [Collections.Generic.List[object]]::new()
$batch = [Collections.Generic.List[object]]::new()
$size = 0
foreach ($exchange in $exchanges) {
    $length = $exchange.Request.Length + $exchange.Response.Length
    if ($batch.Count -gt 0 -and $size + $length -gt $BatchMegabytes * 1MB) {
        $batches.Add($batch)
        $batch = [Collections.Generic.List[object]]::new()
        $size = 0
    }
    $batch.Add($exchange)
    $size += $length
}
$batches.Add($batch)

$handler = [Net.Http.HttpClientHandler]::new()
if ($SkipCertificateCheck) {
    $handler.ServerCertificateCustomValidationCallback = [Net.Http.HttpClientHandler]::DangerousAcceptAnyServerCertificateValidator
}
$client = [Net.Http.HttpClient]::new($handler)
$client.Timeout = [TimeSpan]::FromMinutes(10)
$instance = [Guid]::NewGuid().ToString()

for ($index = 0; $index -lt $batches.Count; $index++) {
    $items = $batches[$index]
    $final = $index -eq $batches.Count - 1
    # all that started before the watermark is in this batch or before: the start of the next batch's first
    $watermark = if ($final) { $items[-1].Started.AddSeconds(1) } else { $batches[$index + 1][0].Started }
    while ($true) {
        $content = [Net.Http.MultipartContent]::new('mixed')
        foreach ($item in $items) {
            $bytes = [byte[]]::new($item.Request.Length + $item.Response.Length)
            [Array]::Copy($item.Request, $bytes, $item.Request.Length)
            [Array]::Copy($item.Response, 0, $bytes, $item.Request.Length, $item.Response.Length)
            $part = [Net.Http.ByteArrayContent]::new($bytes)
            [void] $part.Headers.TryAddWithoutValidation('X-Shadow-Sequence', "$($item.Number)")
            [void] $part.Headers.TryAddWithoutValidation('X-Shadow-Started', $item.Started.UtcDateTime.ToString('o'))
            [void] $part.Headers.TryAddWithoutValidation('X-Shadow-Kind', 'exchange')
            [void] $part.Headers.TryAddWithoutValidation('X-Shadow-Request-Length', "$($item.Request.Length)")
            $content.Add($part)
        }
        $message = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::Post, $Url)
        $message.Content = $content
        [void] $message.Headers.TryAddWithoutValidation('X-OfficeHost-Shadow-Key', $ApiKey)
        [void] $message.Headers.TryAddWithoutValidation('X-Shadow-Instance', $instance)
        [void] $message.Headers.TryAddWithoutValidation('X-Shadow-Server', $Server)
        [void] $message.Headers.TryAddWithoutValidation('X-Shadow-Watermark', $watermark.UtcDateTime.ToString('o'))
        [void] $message.Headers.TryAddWithoutValidation('X-Shadow-Dropped', '0')
        if ($final) {
            [void] $message.Headers.TryAddWithoutValidation('X-Shadow-Final', 'true')
        }
        $response = $client.SendAsync($message).GetAwaiter().GetResult()
        $status = [int] $response.StatusCode
        if ($status -eq 503) {
            $wait = if ($response.Headers.RetryAfter.Delta) { $response.Headers.RetryAfter.Delta.Value } else { [TimeSpan]::FromSeconds(30) }
            Write-Host "The shadow is full: again in $([int] $wait.TotalSeconds) s"
            Start-Sleep -Seconds ([int] $wait.TotalSeconds)
            continue
        }
        if ($status -ne 202) {
            throw "The shadow answered $status to batch $($index + 1)"
        }
        Write-Host "Batch $($index + 1) of $($batches.Count): exchanges $($items[0].Number) to $($items[-1].Number), $status"
        break
    }
}
