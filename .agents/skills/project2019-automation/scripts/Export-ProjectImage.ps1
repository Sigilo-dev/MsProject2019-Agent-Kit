[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ProjectPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateSet('Table','TableAndGantt')][string]$Mode = 'TableAndGantt',
    [string]$OptionsPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot\ProjectCom.ps1"
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use Windows PowerShell 5.1 with -STA.' }
$watch = [Diagnostics.Stopwatch]::StartNew()
$ProjectPath = (Resolve-Path -LiteralPath $ProjectPath).ProviderPath
$OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
if ([IO.Path]::GetExtension($ProjectPath) -ine '.mpp') { throw 'ProjectPath must be a saved local MPP file.' }
if (Test-Path -LiteralPath $OutputDirectory) { throw 'OutputDirectory must be new; existing results are never overwritten.' }
$options = if ($OptionsPath) { Get-Content -LiteralPath $OptionsPath -Raw -Encoding UTF8 | ConvertFrom-Json } else { [pscustomobject]@{} }
foreach ($property in $options.PSObject.Properties) { if ($property.Name -notin @('taskUniqueIds','columns','fromDate','toDate')) { throw "Unsupported image option: $($property.Name)" } }
$fieldIds = @{ID=188743703;UniqueID=188743766;Name=188743694;Duration=188743709;Start=188743715;Finish=188743716;Predecessors=188743727;PercentComplete=188743712;ResourceNames=188743729}
$columns = @(
    [pscustomobject]@{field='ID';width=5},[pscustomobject]@{field='Name';width=38},
    [pscustomobject]@{field='Duration';width=12},[pscustomobject]@{field='Start';width=18},
    [pscustomobject]@{field='Finish';width=18},[pscustomobject]@{field='Predecessors';width=14}
)
if ($options.PSObject.Properties['columns']) { $columns = @($options.columns) }
if ($columns.Count -lt 1 -or $columns.Count -gt 12) { throw 'Choose between 1 and 12 columns.' }
foreach ($col in $columns) {
    foreach ($key in $col.PSObject.Properties.Name) { if ($key -notin @('field','width','title')) { throw "Unsupported column option: $key" } }
    if (-not $fieldIds.ContainsKey([string]$col.field)) { throw "Unsupported field: $($col.field)" }
    if ($null -eq $col.width -or $col.width -is [bool] -or $col.width -is [string] -or $col.width -lt 3 -or $col.width -gt 100 -or $col.width -ne [math]::Floor($col.width)) { throw 'Column width must be a JSON integer between 3 and 100.' }
}
if (@($columns.field | Select-Object -Unique).Count -ne $columns.Count) { throw 'Duplicate columns are not supported.' }
$wantedIds = @()
if ($options.PSObject.Properties['taskUniqueIds']) {
    foreach ($id in $options.taskUniqueIds) {
        if ($null -eq $id -or $id -is [bool] -or $id -is [string] -or $id -lt 1 -or $id -gt [int]::MaxValue -or $id -ne [math]::Floor($id)) { throw 'taskUniqueIds must contain positive JSON integers.' }
        $wantedIds += [int]$id
    }
    if (-not $wantedIds.Count -or @($wantedIds | Select-Object -Unique).Count -ne $wantedIds.Count) { throw 'taskUniqueIds must be nonempty and unique when supplied.' }
}
$from = $null; $to = $null
foreach ($key in @('fromDate','toDate')) {
    if ($options.PSObject.Properties[$key]) {
        if ([string]$options.$key -notmatch '^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$') { throw 'Use local ISO dates for fromDate/toDate.' }
        $date = [datetime]::Parse($options.$key,[Globalization.CultureInfo]::InvariantCulture)
        if ($key -eq 'fromDate') {$from=$date} else {$to=$date}
    }
}
if (($null -eq $from) -ne ($null -eq $to) -or ($from -and $to -le $from)) { throw 'Supply both fromDate and toDate with a positive interval, or neither.' }
if ($Mode -eq 'Table' -and $from) { throw 'Dates apply only to TableAndGantt.' }
if (@(Get-Process -Name WINPROJ -ErrorAction SilentlyContinue).Count -ne 1) { throw 'Expected one Project application instance with an active document.' }
$app=$null; $original=$null; $render=$null; $opened=$false; $copyPath=$null
function Assert-RenderActive {
    if ((Get-ProjectProperty (Get-ProjectProperty $app 'ActiveProject') 'FullName') -ine $copyPath) { throw 'Active document changed; refusing to act on another file.' }
}
function Read-SelectedIds {
    $selection = Get-ProjectProperty $app 'ActiveSelection'
    $tasks = Get-ProjectProperty $selection 'Tasks'
    $rows = @()
    for ($i=1;$i -le (Get-ProjectProperty $tasks 'Count');$i++) {
        $task = $tasks.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$tasks,@($i))
        if ($null -eq $task) { continue }
        $uid = [int](Get-ProjectProperty $task 'UniqueID')
        if ($uid -eq 0) { continue }
        $rows += $uid
    }
    return $rows
}
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application')
    $original = Get-ProjectProperty $app 'ActiveProject'
    $originalName = [string](Get-ProjectProperty $original 'FullName')
    $originalSaved = Get-ProjectProperty $original 'Saved'
    $projects = Get-ProjectProperty $app 'Projects'
    for ($i=1;$i -le (Get-ProjectProperty $projects 'Count');$i++) {
        $candidate = $projects.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$projects,@($i))
        if ((Get-ProjectProperty $candidate 'FullName') -ieq $ProjectPath -and -not (Get-ProjectProperty $candidate 'Saved')) { throw 'Source has unsaved state. Preserve the intended version before exporting a disk copy.' }
    }
    $sourceHash = (Get-FileHash -LiteralPath $ProjectPath -Algorithm SHA256).Hash
    [void][IO.Directory]::CreateDirectory($OutputDirectory)
    $copyPath = Join-Path $OutputDirectory ('render-'+[guid]::NewGuid().ToString('N')+'.mpp')
    [IO.File]::Copy($ProjectPath,$copyPath,$false)
    $copyPath = (Resolve-Path -LiteralPath $copyPath).ProviderPath
    # Read-only local copy, no Auto_Open macro; leave it on disk for diagnosis.
    $m=[Type]::Missing
    if (-not (Invoke-ProjectMethod $app 'FileOpenEx' @($copyPath,$true,$m,$m,$m,$m,$true,$m,$m,$m,$m,$m,$m,$m,$m,$m,$true))) { throw 'Could not open rendering copy.' }
    $render=Get-ProjectProperty $app 'ActiveProject'; $opened=$true
    Assert-RenderActive
    $snapshot=Get-ProjectSnapshot $render
    $view = if ($Mode -eq 'Table') {23} else {4}
    if (-not (Invoke-ProjectMethod $app 'ViewApplyEx' @($m,$true,$false,0,$view))) { throw 'Could not apply task view.' }
    [void](Invoke-ProjectMethod $app 'FilterClear')
    [void](Invoke-ProjectMethod $app 'GroupClear')
    [void](Invoke-ProjectMethod $app 'OutlineShowAllTasks')
    $idName = Invoke-ProjectMethod $app 'FieldConstantToFieldName' @(188743703)
    [void](Invoke-ProjectMethod $app 'Sort' @($idName,$true,$m,$m,$m,$m,$false,$false))
    $tables = Get-ProjectProperty $render 'TaskTables'
    $tableName = 'AgentImage-'+[guid]::NewGuid().ToString('N')
    $table = Invoke-ProjectMethod $tables 'Add' @($tableName,$fieldIds[[string]$columns[0].field],$true)
    $fields = Get-ProjectProperty $table 'TableFields'
    $first = $fields.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$fields,@(1))
    Set-ProjectProperty $first 'Width' ([int]$columns[0].width)
    if ($columns[0].PSObject.Properties['title']) {Set-ProjectProperty $first 'Title' ([string]$columns[0].title)}
    for ($i=1;$i -lt $columns.Count;$i++) {
        $title = if ($columns[$i].PSObject.Properties['title']) {[string]$columns[$i].title} else {$m}
        [void](Invoke-ProjectMethod $fields 'Add' @($fieldIds[[string]$columns[$i].field],0,[int]$columns[$i].width,$title,0,-1,$true))
    }
    [void](Invoke-ProjectMethod $app 'TableApply' @($tableName))
    if ($Mode -eq 'TableAndGantt') { [void](Invoke-ProjectMethod $app 'SetSplitBar' @($columns.Count)) }
    [void](Invoke-ProjectMethod $app 'SelectAll')
    $all = @($snapshot.Tasks | Sort-Object ID)
    $selection=Get-ProjectProperty $app 'ActiveSelection'
    $selectionTasks=Get-ProjectProperty $selection 'Tasks'
    if ((Get-ProjectProperty $selectionTasks 'Count') -ne $all.Count) { throw 'Selection does not contain every task. Inspect filters and outline.' }
    $projectTasks = Get-ProjectProperty $render 'Tasks'
    if (-not $all.Count -or $all.Count -ne (Get-ProjectProperty $projectTasks 'Count')) { throw 'Not all task rows are selectable (blank rows, filters or special project). Inspect the rendering copy.' }
    $selected = $all
    if ($wantedIds.Count) {
        $desired = @($all | Where-Object {$_.UniqueID -in $wantedIds})
        if ($desired.Count -ne $wantedIds.Count) { throw 'One or more requested UniqueIDs do not exist.' }
        $add=$false
        foreach ($row in $desired) {
            [void](Invoke-ProjectMethod $app 'SelectRow' @($row.ID,$false,0,$false,$add))
            $add=$true
        }
        $actualIds = @(Read-SelectedIds)
        if ((@($actualIds | Sort-Object) -join ',') -ne (@($wantedIds | Sort-Object) -join ',')) { throw 'Selected row identities do not match the request.' }
        $selected=$desired
    }
    $rangeFrom=$m; $rangeTo=$m
    if ($Mode -eq 'TableAndGantt') {
        foreach ($row in $selected) { if ($row.Start -eq 'NA' -or $row.Finish -eq 'NA') { throw 'Unscheduled tasks require valid dates before Gantt export.' } }
        $rangeFrom=([datetime]($selected.Start | Sort-Object | Select-Object -First 1)).Date.AddDays(-1)
        $rangeTo=([datetime]($selected.Finish | Sort-Object | Select-Object -Last 1)).Date.AddDays(1)
        if ($from) {$rangeFrom=$from; $rangeTo=$to}
        [void](Invoke-ProjectMethod $app 'ZoomTimescale' @($m,$false,$true,$false))
    }
    Assert-RenderActive
    [string]$gifPath=Join-Path $OutputDirectory 'project.gif'
    # pjGIF=2, SelectedRows=true, KeepRange=1: include off-screen rows, no forced shrinking/truncation.
    if (-not (Invoke-ProjectMethod $app 'EditCopyPicture' @($false,2,$true,$rangeFrom,$rangeTo,$gifPath,1))) { throw 'Native GIF export failed.' }
    $conversion = & "$PSScriptRoot\Convert-ProjectGifToPng.ps1" -InputPath $gifPath -OutputPath (Join-Path $OutputDirectory 'project.png') | ConvertFrom-Json
    if ($conversion.Height -lt ($selected.Count*8)) { throw 'Image is unexpectedly short; possible native truncation. Inspect the GIF.' }
    Assert-RenderActive
    if (-not (Invoke-ProjectMethod $app 'FileCloseEx' @(0,$true))) { throw 'Could not close rendering copy.' }
    $opened=$false
    [void](Invoke-ProjectMethod $original 'Activate')
    if ((Get-ProjectProperty $original 'Saved') -ne $originalSaved) { throw 'Original document state changed during export.' }
    if ((Get-FileHash -LiteralPath $ProjectPath -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Source hash changed.' }
    $report = [pscustomobject]@{
        Result='EXPORTED';Mode=$Mode;Source=$ProjectPath;SourceSHA256=$sourceHash;SourceFileUnchanged=$true
        OriginalRestored=$originalName;RenderingCopy=$copyPath;Columns=$columns
        Rows=$selected.Count;SelectedUniqueIDs=@($selected.UniqueID);SelectionVerified=$true
        FirstTask=$selected[0].Name;LastTask=$selected[-1].Name
        FromDate=$(if($Mode -eq 'TableAndGantt'){$rangeFrom.ToString('o')}else{$null})
        ToDate=$(if($Mode -eq 'TableAndGantt'){$rangeTo.ToString('o')}else{$null})
        Gif=$gifPath;Png=$conversion.Output;Width=$conversion.Width;Height=$conversion.Height
        PixelIdentical=$conversion.PixelIdentical;DecodedArgbSHA256=$conversion.DecodedArgbSHA256
        VisualReviewRequired=$true;ElapsedSeconds=[math]::Round($watch.Elapsed.TotalSeconds,2)
    }
    $json = $report | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'export.json'),$json,[Text.UTF8Encoding]::new($false))
    $json
} finally {
    if ($opened -and $render) {
        try { Assert-RenderActive; [void](Invoke-ProjectMethod $app 'FileCloseEx' @(0,$true)) } catch { Write-Warning 'Rendering copy may still be open; inspect before another operation.' }
    }
    if ($original) { try {[void](Invoke-ProjectMethod $original 'Activate')} catch {Write-Warning 'Could not restore original document.'} }
    if ($app) {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
