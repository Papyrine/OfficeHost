<#
.SYNOPSIS
Summarizes a Cobertura coverage report of the server: line and branch coverage per file, and in total.

.DESCRIPTION
With -Summary, the table is also appended to that file as markdown (the GitHub job summary in CI).
Fails when the report covers no lines, which means the coverage collection did not work.

.EXAMPLE
dotnet run --project src/WopiServer.Tests -- --coverage --coverage-output-format cobertura --coverage-output coverage.cobertura.xml --coverage-settings src/coverage.config --results-directory TestResults
pwsh src/coverage-summary.ps1 TestResults/coverage.cobertura.xml
#>
param(
    [Parameter(Mandatory)] [string] $Report,
    [string] $Summary
)

$ErrorActionPreference = 'Stop'
$xml = [xml] (Get-Content $Report -Raw)
$files = @{}
foreach ($class in $xml.SelectNodes("//package[@name='wopiserver']//class")) {
    $name = ($class.filename -replace '\\', '/') -replace '^.*src/WopiServer/', ''
    if (-not $files.ContainsKey($name)) {
        $files[$name] = @{ Covered = [System.Collections.Generic.HashSet[int]]::new(); All = [System.Collections.Generic.HashSet[int]]::new(); Branches = 0; BranchesCovered = 0 }
    }
    $file = $files[$name]
    foreach ($line in $class.SelectNodes('lines/line')) {
        $number = [int] $line.number
        [void] $file.All.Add($number)
        if ([int] $line.hits -gt 0) {
            [void] $file.Covered.Add($number)
        }
        if ($line.'condition-coverage' -match '\((\d+)/(\d+)\)') {
            $file.BranchesCovered += [int] $Matches[1]
            $file.Branches += [int] $Matches[2]
        }
    }
}

function Percent([double] $part, [double] $whole) {
    if ($whole -eq 0) { return 100 }
    return [math]::Round($part * 100 / $whole, 1)
}

$rows = foreach ($name in $files.Keys) {
    $file = $files[$name]
    [pscustomobject] @{
        File = $name
        Lines = Percent $file.Covered.Count $file.All.Count
        Branches = Percent $file.BranchesCovered $file.Branches
        Covered = $file.Covered.Count
        Total = $file.All.Count
        BranchesCovered = $file.BranchesCovered
        BranchesTotal = $file.Branches
    }
}
$rows = $rows | Sort-Object Lines, File
$lines = ($rows | Measure-Object Covered -Sum).Sum
$total = ($rows | Measure-Object Total -Sum).Sum
$branches = ($rows | Measure-Object BranchesCovered -Sum).Sum
$branchesTotal = ($rows | Measure-Object BranchesTotal -Sum).Sum
if (-not $total) {
    throw "The report $Report covers no line of the server: the coverage collection did not work"
}

$rows | Format-Table File, @{ Label = 'Lines %'; Expression = { $_.Lines } }, @{ Label = 'Branches %'; Expression = { $_.Branches } }, Covered, Total -AutoSize | Out-String | Write-Host
$totalLine = "Total: $(Percent $lines $total)% of lines ($lines/$total), $(Percent $branches $branchesTotal)% of branches ($branches/$branchesTotal)"
Write-Host $totalLine

if ($Summary) {
    $markdown = @('## Coverage', '', $totalLine, '', '| File | Lines | Branches |', '|---|---|---|')
    $markdown += $rows | ForEach-Object { "| $($_.File) | $($_.Lines)% | $($_.Branches)% |" }
    Add-Content $Summary ($markdown -join "`n")
}
