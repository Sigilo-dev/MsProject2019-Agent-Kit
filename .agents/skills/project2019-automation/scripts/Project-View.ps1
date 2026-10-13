param(
    [Parameter(Mandatory=$true)][ValidateSet('Open','Restore')][string]$Action,
    [string]$ProjectPath,
    [Parameter(Mandatory=$true)][string]$SessionPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot\ProjectCom.ps1"
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use Windows PowerShell 5.1 with -STA.' }
$SessionPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SessionPath)
if ([IO.Path]::GetExtension($SessionPath) -ine '.json') { throw 'SessionPath must be a JSON file.' }
$running = @(Get-Process -Name WINPROJ -ErrorAction SilentlyContinue)
if ($running.Count -ne 1) { throw 'Expected one Project application instance.' }
function Find-OpenProject([string]$FullName) {
    $projects = Get-ProjectProperty $app 'Projects'
    for ($i=1;$i -le (Get-ProjectProperty $projects 'Count');$i++) {
        $candidate = $projects.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$projects,@($i))
        if ((Get-ProjectProperty $candidate 'FullName') -ieq $FullName) { return ,$candidate }
    }
    return $null
}
$app = $null
try {
    if ($Action -eq 'Open') {
        if (Test-Path -LiteralPath $SessionPath) { throw 'Session already exists; use a new path or restore the existing session.' }
        if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($SessionPath)) -PathType Container)) { throw 'Session folder does not exist.' }
        $ProjectPath = (Resolve-Path -LiteralPath $ProjectPath).ProviderPath
        if ([IO.Path]::GetExtension($ProjectPath) -ine '.mpp') { throw 'Expected a local MPP file.' }
    } else {
        $session = Get-Content -LiteralPath $SessionPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($session.Kind -ne 'Project2019ViewSession-v1' -or $session.Status -ne 'Open') { throw 'Session is not an open Project view session.' }
        if ($session.ProcessId -ne $running[0].Id) { throw 'Project application instance changed. Inspect it manually before restoring.' }
    }
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application')
    if ($Action -eq 'Open') {
        $original = Get-ProjectProperty $app 'ActiveProject'
        $target = Find-OpenProject $ProjectPath
        $openedByHelper = $null -eq $target
        $session = [ordered]@{
            Kind='Project2019ViewSession-v1';Status='Opening';ProcessId=$running[0].Id
            OriginalFullName=Get-ProjectProperty $original 'FullName'
            TargetFullName=$ProjectPath;OpenedByHelper=$openedByHelper
            SourceSHA256=(Get-FileHash -LiteralPath $ProjectPath -Algorithm SHA256).Hash
        }
        # Reserve a journal before opening; a failed open remains inspectable.
        $stream = [IO.File]::Open($SessionPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        try { $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($session | ConvertTo-Json)); $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
        if ($openedByHelper) {
            $m = [Type]::Missing
            if (-not (Invoke-ProjectMethod $app 'FileOpenEx' @($ProjectPath,$true,$m,$m,$m,$m,$true,$m,$m,$m,$m,$m,$m,$m,$m,$m,$true))) { throw 'Open failed. Inspect the Opening journal before retrying.' }
            $target = Get-ProjectProperty $app 'ActiveProject'
            if ((Get-ProjectProperty $target 'FullName') -ine $ProjectPath) { throw 'Unexpected opened project.' }
        }
        [void](Invoke-ProjectMethod $target 'Activate')
        $session.Status = 'Open'
        [IO.File]::WriteAllText($SessionPath,($session | ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        [pscustomobject]@{Result='OPEN';Project=$ProjectPath;OpenedReadOnlyByHelper=$openedByHelper;Session=$SessionPath;Next='Activate the actual Project window, capture PNG, inspect it, then Restore.'} | ConvertTo-Json
    } else {
        if ((Get-FileHash -LiteralPath $session.TargetFullName -Algorithm SHA256).Hash -ne $session.SourceSHA256) { throw 'Viewed file changed; inspect before restoring.' }
        $target = Find-OpenProject $session.TargetFullName
        $original = Find-OpenProject $session.OriginalFullName
        if ($session.OpenedByHelper -and $target) {
            [void](Invoke-ProjectMethod $target 'Activate')
            $active = Get-ProjectProperty $app 'ActiveProject'
            if ((Get-ProjectProperty $active 'FullName') -ine $session.TargetFullName) { throw 'Active document changed; not closing anything.' }
            if (-not (Invoke-ProjectMethod $app 'FileCloseEx' @(0,$true))) { throw 'Could not close the helper-opened view.' }
        }
        if (-not $original) { throw 'Original document is no longer open. No other document was opened as a substitute.' }
        [void](Invoke-ProjectMethod $original 'Activate')
        $session.Status = 'Restored'
        [IO.File]::WriteAllText($SessionPath,($session | ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        [pscustomobject]@{Result='RESTORED';FileUnchanged=$true;Original=$session.OriginalFullName} | ConvertTo-Json
    }
} finally { if ($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
