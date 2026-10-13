$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. "$PSScriptRoot\ProjectCom.ps1"
if ($PSVersionTable.PSEdition -ne 'Desktop' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use Windows PowerShell 5.1 with -STA.' }
$running = @(Get-Process -Name WINPROJ -ErrorAction SilentlyContinue)
if ($running.Count -ne 1) { throw 'Expected exactly one running Project application. Open Project with a blank or intended document; do not close user documents automatically.' }
$app = $null
try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application')
    $projects = Get-ProjectProperty $app 'Projects'
    $active = Get-ProjectProperty $app 'ActiveProject'
    $tasks = Get-ProjectProperty $active 'Tasks'
    [pscustomobject]@{
        Result='READY';PowerShell=$PSVersionTable.PSVersion.ToString()
        Apartment=[Threading.Thread]::CurrentThread.ApartmentState.ToString()
        Is64BitProcess=[Environment]::Is64BitProcess
        ApplicationVersion=Get-ProjectProperty $app 'Version'
        OpenProjects=Get-ProjectProperty $projects 'Count'
        ActiveProject=Get-ProjectProperty $active 'FullName'
        Saved=Get-ProjectProperty $active 'Saved'
        TaskCount=Get-ProjectProperty $tasks 'Count'
    } | ConvertTo-Json
} finally { if ($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
