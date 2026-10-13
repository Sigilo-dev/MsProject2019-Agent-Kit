param([Parameter(Mandatory=$true)][string]$ScratchPath)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$ScratchPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ScratchPath)
if (Test-Path -LiteralPath $ScratchPath) { throw 'Use a new scratch folder.' }
[void][IO.Directory]::CreateDirectory($ScratchPath)
$valid = @{tasks=@(@{key='a';name='Task A';durationMinutes=480})}
$cases = @(
    @{Name='report-collision';Plan=$valid;Error='ReportPath must end with';Collision=$true},
    @{Name='cycle';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=480;predecessorKeys=@('b')},@{key='b';name='B';durationMinutes=480;predecessorKeys=@('a')})};Error='Dependency cycle';Collision=$false},
    @{Name='fractional-percent';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=480;percentComplete=2.5})};Error='integer between';Collision=$false},
    @{Name='unknown-field';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=480;unrecognized='x'})};Error='Unsupported task field';Collision=$false},
    @{Name='null-duration';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=$null})};Error='durationMinutes must be a JSON number';Collision=$false},
    @{Name='boolean-duration';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=$true})};Error='durationMinutes must be a JSON number';Collision=$false},
    @{Name='string-duration';Plan=@{tasks=@(@{key='a';name='A';durationMinutes='480'})};Error='durationMinutes must be a JSON number';Collision=$false},
    @{Name='null-percent';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=480;percentComplete=$null})};Error='percentComplete must be a JSON number';Collision=$false},
    @{Name='boolean-percent';Plan=@{tasks=@(@{key='a';name='A';durationMinutes=480;percentComplete=$false})};Error='percentComplete must be a JSON number';Collision=$false}
)
$results = @()
foreach ($case in $cases) {
    $planPath = Join-Path $ScratchPath ($case.Name+'.json')
    $outputPath = Join-Path $ScratchPath ($case.Name+'.mpp')
    $reportPath = if ($case.Collision) { $outputPath } else { Join-Path $ScratchPath ($case.Name+'-report.json') }
    [IO.File]::WriteAllText($planPath,($case.Plan | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $messages = & powershell.exe -NoProfile -NonInteractive -STA -File "$PSScriptRoot\Project-Codex.ps1" -Action Create -PlanPath $planPath -OutputPath $outputPath -ReportPath $reportPath 2>&1
    $code = $LASTEXITCODE
    $ErrorActionPreference = $oldPreference
    # Inspect raw error text; the default formatter can wrap a phrase across lines.
    $messageText = (($messages | ForEach-Object { [string]$_ }) -join ' ') -replace '\s+',' '
    if ($code -eq 0 -or ($messageText -notmatch [regex]::Escape($case.Error)) -or (Test-Path -LiteralPath $outputPath)) { throw "Guard failed: $($case.Name). Exit=$code. $messageText" }
    $results += [pscustomobject]@{Name=$case.Name;Result='PASS';ProjectCreated=$false}
}
$results | ConvertTo-Json
