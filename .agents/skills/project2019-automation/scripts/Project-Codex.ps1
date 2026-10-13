[CmdletBinding()]
param(
    [ValidateSet('Inspect','Create','UpdateCopy')][string]$Action = 'Inspect',
    [string]$PlanPath,
    [string]$OutputPath,
    [string]$SourcePath,
    [string]$ReportPath,
    [string]$ExpectedActiveProjectPath
)
# Run with Windows PowerShell 5.1 -STA. Uses Project COM, not screen automation.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot\ProjectCom.ps1"
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Use powershell.exe -NoProfile -NonInteractive -STA -File Project-Codex.ps1 ...'
}
function Assert-ActiveProject([object]$Expected) {
    $active = Get-ProjectProperty $app 'ActiveProject'
    if ((Get-ProjectProperty $active 'FullName') -ne (Get-ProjectProperty $Expected 'FullName')) {
        throw 'The active project changed. Stopping before saving or closing.'
    }
}
function Assert-EqualSnapshot($Expected,$Actual,[switch]$PersistedOnly) {
    if ($PersistedOnly) {
        # Name extension and Saved are session metadata, not persisted task data.
        $Expected = $Expected | Select-Object FullName,Start,HoursPerDay,MinutesPerDayConversion,TaskCount,Tasks
        $Actual = $Actual | Select-Object FullName,Start,HoursPerDay,MinutesPerDayConversion,TaskCount,Tasks
    }
    $expectedText = $Expected | ConvertTo-Json -Depth 12 -Compress
    $actualText = $Actual | ConvertTo-Json -Depth 12 -Compress
    if ($expectedText -cne $actualText) { throw 'Saved and reopened project data did not match.' }
}
function Set-TaskFields($Task,$Spec) {
    if ($Spec.PSObject.Properties['name']) { Set-ProjectProperty $Task 'Name' ([string]$Spec.name) }
    if ($Spec.PSObject.Properties['durationMinutes']) { Set-ProjectProperty $Task 'Duration' ([double]$Spec.durationMinutes) }
    if ($Spec.PSObject.Properties['percentComplete']) { Set-ProjectProperty $Task 'PercentComplete' ([int]$Spec.percentComplete) }
}
function Assert-JsonNumber($Value,[string]$Field) {
    if ($null -eq $Value -or ($Value -isnot [int] -and $Value -isnot [long] -and $Value -isnot [double] -and $Value -isnot [decimal])) {
        throw "$Field must be a JSON number; null, booleans and strings are not accepted."
    }
}
function Assert-TaskFields($Spec) {
    if ($Spec.PSObject.Properties['name'] -and [string]::IsNullOrWhiteSpace([string]$Spec.name)) { throw 'Task name must be nonempty.' }
    if ($Spec.PSObject.Properties['durationMinutes']) {
        Assert-JsonNumber $Spec.durationMinutes 'durationMinutes'
        $duration = [double]$Spec.durationMinutes
        if ([double]::IsNaN($duration) -or [double]::IsInfinity($duration) -or $duration -lt 0) { throw 'Duration must be a finite nonnegative number.' }
    }
    if ($Spec.PSObject.Properties['percentComplete']) {
        Assert-JsonNumber $Spec.percentComplete 'percentComplete'
        $percent = [double]$Spec.percentComplete
        if ([double]::IsNaN($percent) -or $percent -ne [math]::Floor($percent) -or $percent -lt 0 -or $percent -gt 100) { throw 'Percent complete must be an integer between 0 and 100.' }
    }
}
function Assert-RequestedFields($Spec,$Row) {
    if ($Spec.PSObject.Properties['name'] -and $Row.Name -cne [string]$Spec.name) { throw 'Requested task name was not applied.' }
    if ($Spec.PSObject.Properties['durationMinutes'] -and [double]$Row.Duration -ne [double]$Spec.durationMinutes) { throw 'Requested duration was not applied.' }
    if ($Spec.PSObject.Properties['percentComplete'] -and [int]$Row.PercentComplete -ne [int]$Spec.percentComplete) { throw 'Requested percent complete was not applied.' }
}
function Open-LocalProject([string]$Path,[bool]$ReadOnly) {
    $m = [Type]::Missing
    $ok = Invoke-ProjectMethod $app 'FileOpenEx' @($Path,$ReadOnly,$m,$m,$m,$m,$true,$m,$m,$m,$m,$m,$m,$m,$m,$m,$true)
    if (-not $ok) { throw "Project could not open $Path" }
    $opened = Get-ProjectProperty $app 'ActiveProject'
    if ((Get-ProjectProperty $opened 'FullName') -ne $Path) { throw 'Unexpected project opened.' }
    return ,$opened
}
$app = $null
$original = $null
$sourceHash = $null
try {
    if ($ReportPath) {
        $ReportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
        if ([IO.Path]::GetExtension($ReportPath) -ine '.json') { throw 'ReportPath must end with .json.' }
        if (Test-Path -LiteralPath $ReportPath) { throw 'Report already exists. Choose a new path.' }
        if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($ReportPath)) -PathType Container)) { throw 'Report folder does not exist.' }
    }
    if ($Action -ne 'Inspect') {
        if (-not $PlanPath -or -not $OutputPath) { throw 'PlanPath and OutputPath are required.' }
        $plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        if ([IO.Path]::GetExtension($OutputPath) -ine '.mpp') { throw 'OutputPath must be a new .mpp file.' }
        if (Test-Path -LiteralPath $OutputPath) { throw 'Output already exists. Choose a new path; overwriting is disabled.' }
        if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($OutputPath)) -PathType Container)) { throw 'Output folder does not exist.' }
        if ($Action -eq 'Create') {
            $keys = @{}
            if (@($plan.tasks).Count -eq 0) { throw 'At least one task is required.' }
            if ($plan.PSObject.Properties['startDate']) { [void][datetime]::Parse($plan.startDate,[Globalization.CultureInfo]::InvariantCulture) }
            foreach ($spec in $plan.tasks) {
                foreach ($field in $spec.PSObject.Properties.Name) { if ($field -notin @('key','name','durationMinutes','percentComplete','predecessorKeys')) { throw "Unsupported task field: $field" } }
                if (-not $spec.key -or $keys.ContainsKey([string]$spec.key) -or -not $spec.name) { throw 'Task keys must be unique and names nonempty.' }
                if (-not $spec.PSObject.Properties['durationMinutes']) { throw 'Every task requires durationMinutes.' }
                $keys[[string]$spec.key] = $true
                Assert-TaskFields $spec
            }
            $pending = @($plan.tasks)
            $resolved = @{}
            while ($pending.Count -gt 0) {
                $remaining = @()
                foreach ($spec in $pending) {
                    $deps = @()
                    if ($spec.PSObject.Properties['predecessorKeys']) { $deps = @($spec.predecessorKeys) }
                    if (@($deps | Where-Object {-not $resolved.ContainsKey([string]$_)}).Count -eq 0) { $resolved[[string]$spec.key] = $true }
                    else { $remaining += $spec }
                }
                if ($remaining.Count -eq $pending.Count) { throw 'Dependency cycle detected; no project was created.' }
                $pending = $remaining
            }
            foreach ($spec in $plan.tasks) {
                if ($spec.PSObject.Properties['predecessorKeys']) {
                    foreach ($key in $spec.predecessorKeys) { if (-not $keys.ContainsKey([string]$key) -or $key -eq $spec.key) { throw 'Invalid predecessor key.' } }
                }
            }
        } else {
            $SourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
            if ([IO.Path]::GetExtension($SourcePath) -ine '.mpp') { throw 'Source must be .mpp.' }
            $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
            if (-not $plan.PSObject.Properties['updates']) { throw 'UpdateCopy requires updates with uniqueId.' }
            if (@($plan.updates).Count -eq 0) { throw 'At least one update is required.' }
            $updateIds = @{}
            foreach ($spec in $plan.updates) {
                foreach ($field in $spec.PSObject.Properties.Name) { if ($field -notin @('uniqueId','name','durationMinutes','percentComplete')) { throw "Unsupported update field: $field" } }
                Assert-JsonNumber $spec.uniqueId 'uniqueId'
                $uniqueId = [double]$spec.uniqueId
                if ([double]::IsNaN($uniqueId) -or [double]::IsInfinity($uniqueId) -or $uniqueId -ne [math]::Floor($uniqueId) -or $uniqueId -lt 1 -or $uniqueId -gt [int]::MaxValue) { throw 'Each update needs a positive integer uniqueId.' }
                if ($updateIds.ContainsKey([int]$uniqueId)) { throw 'Duplicate uniqueId in updates.' }
                $updateIds[[int]$uniqueId] = $true
                Assert-TaskFields $spec
            }
        }
        if ($ReportPath) {
            $ReportPath = [IO.Path]::GetFullPath($ReportPath)
            if ($ReportPath -ieq $OutputPath -or ($SourcePath -and $ReportPath -ieq $SourcePath)) { throw 'Report must be separate from the project files.' }
            if ([IO.Path]::GetExtension($ReportPath) -ine '.json') { throw 'ReportPath must end with .json.' }
            if (Test-Path -LiteralPath $ReportPath) { throw 'Report already exists. Choose a new path.' }
        }
    }
    $runningProjects = @(Get-Process -Name WINPROJ -ErrorAction SilentlyContinue)
    if ($runningProjects.Count -ne 1) { throw 'Open exactly one Project application instance, with an active document. Do not close other instances automatically.' }
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application')
    $original = Get-ProjectProperty $app 'ActiveProject'
    if ($ExpectedActiveProjectPath -and (Get-ProjectProperty $original 'FullName') -ine $ExpectedActiveProjectPath) { throw 'The active project does not match ExpectedActiveProjectPath.' }
    $before = Get-ProjectSnapshot $original
    $version = Get-ProjectProperty $app 'Version'
    if ($Action -eq 'Inspect') {
        $inspectJson = [pscustomobject]@{Version=$version;Project=$before} | ConvertTo-Json -Depth 12
        if ($ReportPath) {
            $inspectStream = [IO.File]::Open($ReportPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
            try { $inspectBytes = [Text.UTF8Encoding]::new($false).GetBytes($inspectJson); $inspectStream.Write($inspectBytes,0,$inspectBytes.Length) } finally { $inspectStream.Dispose() }
        }
        $inspectJson
        return
    }
    $projects = Get-ProjectProperty $app 'Projects'
    if ($Action -eq 'Create') {
        $test = Invoke-ProjectMethod $projects 'Add' @($false,[Type]::Missing,$false)
        if ($plan.PSObject.Properties['startDate']) { Set-ProjectProperty $test 'ProjectStart' ([datetime]::Parse($plan.startDate,[Globalization.CultureInfo]::InvariantCulture)) }
        $tasks = Get-ProjectProperty $test 'Tasks'
        $taskMap = @{}
        foreach ($spec in $plan.tasks) {
            $task = Invoke-ProjectMethod $tasks 'Add' @([string]$spec.name)
            Set-ProjectProperty $task 'Manual' $false
            Set-TaskFields $task $spec
            $taskMap[[string]$spec.key] = $task
        }
        foreach ($spec in $plan.tasks) {
            if ($spec.PSObject.Properties['predecessorKeys']) {
                foreach ($key in $spec.predecessorKeys) {
                    $links = Get-ProjectProperty $taskMap[[string]$spec.key] 'TaskDependencies'
                    [void](Invoke-ProjectMethod $links 'Add' @($taskMap[[string]$key],1,0))
                }
            }
        }
    } else {
        $projectCount = [int](Get-ProjectProperty $projects 'Count')
        for ($i=1;$i -le $projectCount;$i++) {
            $candidate = $projects.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$projects,@($i))
            if ((Get-ProjectProperty $candidate 'FullName') -eq $SourcePath -and -not (Get-ProjectProperty $candidate 'Saved')) { throw 'Source has unsaved edits. Save it before copying.' }
        }
        [IO.File]::Copy($SourcePath,$OutputPath,$false)
        $test = Open-LocalProject $OutputPath $false
        $tasks = Get-ProjectProperty $test 'Tasks'
        foreach ($spec in $plan.updates) {
            $task = $tasks.GetType().InvokeMember('UniqueID',[Reflection.BindingFlags]::GetProperty,$null,$tasks,@([int]$spec.uniqueId))
            if ($null -eq $task) { throw "Task UniqueID $($spec.uniqueId) not found." }
            Set-TaskFields $task $spec
        }
    }
    [void](Invoke-ProjectMethod $test 'Activate')
    Assert-ActiveProject $test
    if ($Action -eq 'Create') { $saved = Invoke-ProjectMethod $app 'FileSaveAs' @($OutputPath) }
    else { $saved = Invoke-ProjectMethod $app 'FileSave' }
    if (-not $saved -or -not (Test-Path -LiteralPath $OutputPath) -or (Get-ProjectProperty $test 'FullName') -ne $OutputPath) { throw 'Save verification failed.' }
    $expected = Get-ProjectSnapshot $test
    if ($Action -eq 'Create') {
        if ($expected.TaskCount -ne @($plan.tasks).Count) { throw 'Unexpected task count.' }
        foreach ($spec in $plan.tasks) {
            $uid = [int](Get-ProjectProperty $taskMap[[string]$spec.key] 'UniqueID')
            $row = @($expected.Tasks | Where-Object {$_.UniqueID -eq $uid})
            if ($row.Count -ne 1 -or $row[0].Manual) { throw 'Task was not created with automatic scheduling.' }
            Assert-RequestedFields $spec $row[0]
            $desiredPredecessors = @()
            if ($spec.PSObject.Properties['predecessorKeys']) {
                foreach ($key in $spec.predecessorKeys) { $desiredPredecessors += [int](Get-ProjectProperty $taskMap[[string]$key] 'UniqueID') }
            }
            if ((@($desiredPredecessors | Sort-Object) -join ',') -ne (@($row[0].PredecessorUniqueIDs) -join ',')) { throw 'Requested predecessors were not applied.' }
        }
    } else {
        foreach ($spec in $plan.updates) {
            $row = @($expected.Tasks | Where-Object {$_.UniqueID -eq [int]$spec.uniqueId})
            if ($row.Count -ne 1) { throw 'Updated task could not be identified.' }
            Assert-RequestedFields $spec $row[0]
        }
    }
    Assert-ActiveProject $test
    if (-not (Invoke-ProjectMethod $app 'FileCloseEx' @(0,$true))) { throw 'Could not close the saved test project.' }
    $test = Open-LocalProject $OutputPath $true
    $actual = Get-ProjectSnapshot $test
    Assert-EqualSnapshot $expected $actual -PersistedOnly
    Assert-ActiveProject $test
    if (-not (Invoke-ProjectMethod $app 'FileCloseEx' @(0,$true))) { throw 'Could not close verification project.' }
    [void](Invoke-ProjectMethod $original 'Activate')
    Assert-EqualSnapshot $before (Get-ProjectSnapshot $original)
    if ($sourceHash -and (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Source file hash changed.' }
    $report = [pscustomobject]@{
        Result='PASS';Action=$Action;Version=$version
        VerifiedUtc=[datetime]::UtcNow.ToString('o');Output=$OutputPath
        ReopenedVerified=$true;RequestedValuesVerified=$true;OriginalSnapshotUnchanged=$true
        SourceFileUnchanged=$(if($sourceHash){$true}else{$null})
        SHA256=(Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash
        Project=$actual
    }
    $json = $report | ConvertTo-Json -Depth 12
    if ($ReportPath) {
        $reportStream = [IO.File]::Open($ReportPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try {
            $bytes = [Text.UTF8Encoding]::new($false).GetBytes($json)
            $reportStream.Write($bytes,0,$bytes.Length)
        } finally { $reportStream.Dispose() }
    }
    $json
} catch {
    $cause = $_.Exception
    while ($cause.InnerException) { $cause = $cause.InnerException }
    throw ('Project operation failed. No automatic retry. HRESULT=0x{0:X8}: {1}' -f $cause.HResult,$cause.Message)
} finally {
    if ($original) { try { [void](Invoke-ProjectMethod $original 'Activate') } catch { Write-Warning 'Could not restore original project focus.' } }
    if ($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
