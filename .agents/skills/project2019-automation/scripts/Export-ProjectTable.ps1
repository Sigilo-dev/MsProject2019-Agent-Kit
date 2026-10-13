param(
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$report = Get-Content -LiteralPath $ReportPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($report.PSObject.Properties['Result'] -and $report.Result -ne 'PASS') { throw 'Cannot export a failed validation report.' }
if (-not $report.PSObject.Properties['Project'] -or -not $report.Project.PSObject.Properties['Tasks']) { throw 'Expected a Project-Codex JSON snapshot.' }
$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$csvPath = Join-Path $OutputDirectory 'tasks.csv'
$mdPath = Join-Path $OutputDirectory 'tasks.md'
foreach ($path in @($csvPath,$mdPath)) { if (Test-Path -LiteralPath $path) { throw "Output already exists: $path" } }
[void][IO.Directory]::CreateDirectory($OutputDirectory)
function Safe-CsvText([object]$Value) {
    $text = [string]$Value
    if ($text -match '^[=+@\-\t\r\n]') { return "'" + $text }
    return $text
}
function Markdown-Cell([object]$Value) {
    $text = [string]$Value
    return $text.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('|','&#124;').Replace("`r`n",'<br>').Replace("`n",'<br>').Replace("`r",'<br>')
}
$rows = @($report.Project.Tasks | ForEach-Object {
    [pscustomobject][ordered]@{
        ID=$_.ID;UniqueID=$_.UniqueID;Task=Safe-CsvText $_.Name
        DurationMinutes=$_.Duration;Start=$_.Start;Finish=$_.Finish
        Predecessors=Safe-CsvText $_.Predecessors;PercentComplete=$_.PercentComplete
    }
})
if ($rows.Count -eq 0) { throw 'Snapshot has no task rows to export.' }
$csv = ($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
$lines = @('# Task table', '', '| ID | UniqueID | Task | Duration (min) | Start | Finish | Predecessors | Complete (%) |', '| --- | --- | --- | --- | --- | --- | --- | --- |')
foreach ($task in $report.Project.Tasks) {
    $cells = @($task.ID,$task.UniqueID,$task.Name,$task.Duration,$task.Start,$task.Finish,$task.Predecessors,$task.PercentComplete) | ForEach-Object { Markdown-Cell $_ }
    $lines += '| ' + ($cells -join ' | ') + ' |'
}
foreach ($entry in @(@{Path=$csvPath;Text=$csv;Bom=$true},@{Path=$mdPath;Text=($lines -join "`n");Bom=$false})) {
    $stream = [IO.File]::Open($entry.Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    $writer = [IO.StreamWriter]::new($stream,[Text.UTF8Encoding]::new($entry.Bom))
    try { $writer.Write($entry.Text) } finally { $writer.Dispose() }
}
[pscustomobject]@{Result='EXPORTED';Rows=$rows.Count;Csv=$csvPath;Markdown=$mdPath} | ConvertTo-Json
