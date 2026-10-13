param(
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$ExpectedTitle = '',
    [long]$WindowHandle = 0
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing,System.Windows.Forms
Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class ProjectScreenRead {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left,Top,Right,Bottom; }
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint id);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
'@
[void][ProjectScreenRead]::SetProcessDPIAware()
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($OutputPath) -ine '.png') { throw 'Output must be PNG.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Output already exists; overwriting is disabled.' }
$foreground = [ProjectScreenRead]::GetForegroundWindow()
$target = if ($WindowHandle) { [IntPtr]$WindowHandle } else { $foreground }
if ($target -ne $foreground) { throw 'Project must be in the foreground before capture.' }
$projectProcessId = [uint32]0
[void][ProjectScreenRead]::GetWindowThreadProcessId($target,[ref]$projectProcessId)
if ((Get-Process -Id $projectProcessId).ProcessName -ine 'WINPROJ') { throw 'Foreground window is not Microsoft Project.' }
if ([ProjectScreenRead]::IsIconic($target)) { throw 'Project is minimized.' }
$titleBuffer = [Text.StringBuilder]::new(1024)
[void][ProjectScreenRead]::GetWindowText($target,$titleBuffer,$titleBuffer.Capacity)
$title = $titleBuffer.ToString()
if ($ExpectedTitle -and -not $title.Contains($ExpectedTitle)) { throw 'Unexpected Project document title.' }
$rect = New-Object ProjectScreenRead+RECT
if (-not [ProjectScreenRead]::GetWindowRect($target,[ref]$rect)) { throw 'Cannot read window bounds.' }
$bounds = [Drawing.Rectangle]::FromLTRB($rect.Left,$rect.Top,$rect.Right,$rect.Bottom)
$bounds = [Drawing.Rectangle]::Intersect($bounds,[Windows.Forms.SystemInformation]::VirtualScreen)
if ($bounds.Width -le 0 -or $bounds.Height -le 0) { throw 'Project is outside the visible desktop.' }
$bitmap = [Drawing.Bitmap]::new($bounds.Width,$bounds.Height)
$graphics = [Drawing.Graphics]::FromImage($bitmap)
try {
    $graphics.CopyFromScreen($bounds.X,$bounds.Y,0,0,$bounds.Size,[Drawing.CopyPixelOperation]::SourceCopy)
    if ([ProjectScreenRead]::GetForegroundWindow() -ne $target) { throw 'Foreground changed during capture; no image saved.' }
    $stream = [IO.File]::Open($OutputPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try { $bitmap.Save($stream,[Drawing.Imaging.ImageFormat]::Png) } finally { $stream.Dispose() }
    [pscustomobject]@{Path=$OutputPath;Title=$title;Width=$bitmap.Width;Height=$bitmap.Height;Method='Windows GDI screen capture'} | ConvertTo-Json
} finally { $graphics.Dispose(); $bitmap.Dispose() }
