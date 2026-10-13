param([Parameter(Mandatory=$true)][string]$RunDirectory)
$ErrorActionPreference='Stop'
$RunDirectory=$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RunDirectory)
if (Test-Path -LiteralPath $RunDirectory) {throw 'Use a new test directory.'}
[void][IO.Directory]::CreateDirectory($RunDirectory)
$rows=@(for($i=1;$i -le 120;$i++) {
    $row=[ordered]@{key="t$i";name=('Task {0:D3} - native export' -f $i);durationMinutes=480}
    if (($i-1)%12 -ne 0) {$row.predecessorKeys=@('t'+($i-1))}
    if ($i -eq 1) {$row.name='START 001 - '+[char]0x00F1+' & <test>'; $row.percentComplete=25}
    if ($i -eq 120) {$row.name='END 120 - milestone';$row.durationMinutes=0}
    [pscustomobject]$row
})
$plan=[pscustomobject]@{startDate='2026-10-12T08:00:00';tasks=$rows}
[IO.File]::WriteAllText((Join-Path $RunDirectory 'plan.json'),($plan|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
$watch=[Diagnostics.Stopwatch]::StartNew()
& powershell.exe -NoProfile -NonInteractive -STA -File "$PSScriptRoot\Project-Codex.ps1" -Action Create -CreationMode Xml -PlanPath "$RunDirectory\plan.json" -OutputPath "$RunDirectory\large.mpp" -ReportPath "$RunDirectory\create.json" > "$RunDirectory\create.log"
if ($LASTEXITCODE -ne 0) {throw 'XML creation failed; inspect create.log.'}
$creationSeconds=$watch.Elapsed.TotalSeconds
$create=Get-Content -LiteralPath "$RunDirectory\create.json" -Raw -Encoding UTF8 | ConvertFrom-Json
if ($create.Result -ne 'PASS' -or $create.Project.TaskCount -ne 120 -or $create.Project.Tasks[0].PercentComplete -ne 25 -or -not $create.Project.Tasks[-1].Milestone) {throw 'Unexpected fixture data.'}
foreach($mode in @('Table','TableAndGantt')) {
    & powershell.exe -NoProfile -NonInteractive -STA -File "$PSScriptRoot\Export-ProjectImage.ps1" -ProjectPath "$RunDirectory\large.mpp" -OutputDirectory "$RunDirectory\$mode" -Mode $mode > "$RunDirectory\$mode.log"
    if ($LASTEXITCODE -ne 0) {throw "Export failed: $mode"}
    $report=Get-Content -LiteralPath "$RunDirectory\$mode\export.json" -Raw | ConvertFrom-Json
    if ($report.Rows -ne 120 -or -not $report.PixelIdentical -or -not $report.SourceFileUnchanged -or $report.Height -le 1080) {throw "Export verification failed: $mode"}
}
$options=@{taskUniqueIds=@(1,60,120);columns=@(@{field='ID';width=5},@{field='Name';width=45},@{field='Duration';width=12},@{field='PercentComplete';width=14;title='Progress'})}
[IO.File]::WriteAllText("$RunDirectory\subset.json",($options|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
& powershell.exe -NoProfile -NonInteractive -STA -File "$PSScriptRoot\Export-ProjectImage.ps1" -ProjectPath "$RunDirectory\large.mpp" -OutputDirectory "$RunDirectory\subset" -Mode Table -OptionsPath "$RunDirectory\subset.json" > "$RunDirectory\subset.log"
if ($LASTEXITCODE -ne 0) {throw 'Subset export failed.'}
$subset=Get-Content -LiteralPath "$RunDirectory\subset\export.json" -Raw | ConvertFrom-Json
if (($subset.SelectedUniqueIDs -join ',') -ne '1,60,120' -or -not $subset.PixelIdentical) {throw 'Wrong subset.'}
[pscustomobject]@{Result='PASS';Rows=120;CreationAndReopenSeconds=[math]::Round($creationSeconds,2);TotalSeconds=[math]::Round($watch.Elapsed.TotalSeconds,2);VisualReviewRequired=$true;Directory=$RunDirectory} | ConvertTo-Json | Tee-Object -FilePath "$RunDirectory\test-result.json"
