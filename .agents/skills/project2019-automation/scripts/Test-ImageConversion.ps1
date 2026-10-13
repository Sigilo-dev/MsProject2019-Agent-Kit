param([Parameter(Mandatory=$true)][string]$RunDirectory)
$ErrorActionPreference='Stop'
$RunDirectory=$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RunDirectory)
if (Test-Path -LiteralPath $RunDirectory) {throw 'Use a new directory.'}
[void][IO.Directory]::CreateDirectory($RunDirectory)
$header='47494638396101000100800000000000ffffff'
$frame='21f90400000000002c0000000001000100000202440100'
function Write-Hex([string]$Path,[string]$Hex) {
    $bytes=[byte[]]::new($Hex.Length/2)
    for ($i=0;$i -lt $bytes.Length;$i++) {$bytes[$i]=[Convert]::ToByte($Hex.Substring($i*2,2),16)}
    [IO.File]::WriteAllBytes($Path,$bytes)
}
Write-Hex "$RunDirectory\static.gif" ($header+$frame+'3b')
Write-Hex "$RunDirectory\animated.gif" ($header+$frame+$frame+'3b')
$converted=& "$PSScriptRoot\Convert-ProjectGifToPng.ps1" -InputPath "$RunDirectory\static.gif" -OutputPath "$RunDirectory\static.png" | ConvertFrom-Json
if (-not $converted.PixelIdentical -or $converted.Width -ne 1 -or $converted.Height -ne 1) {throw 'Lossless conversion failed.'}
[IO.File]::WriteAllText("$RunDirectory\invalid.gif",'not a GIF')
[IO.File]::WriteAllText("$RunDirectory\existing.png",'sentinel')
$cases=@(
    @{Name='animated';Input="$RunDirectory\animated.gif";Output="$RunDirectory\animation.png"},
    @{Name='wrong-format';Input="$RunDirectory\invalid.gif";Output="$RunDirectory\invalid.png"},
    @{Name='overwrite';Input="$RunDirectory\static.gif";Output="$RunDirectory\existing.png"}
)
$results=@([pscustomobject]@{Name='pixel-preservation';Result='PASS'})
foreach($case in $cases) {
    $ErrorActionPreference='Continue'
    $messages=& powershell.exe -NoProfile -NonInteractive -STA -File "$PSScriptRoot\Convert-ProjectGifToPng.ps1" -InputPath $case.Input -OutputPath $case.Output 2>&1
    $code=$LASTEXITCODE
    $ErrorActionPreference='Stop'
    if ($code -eq 0) {throw ('Unexpected conversion success: '+$case.Name)}
    if ($case.Name -eq 'overwrite') {if ([IO.File]::ReadAllText($case.Output) -ne 'sentinel') {throw 'Existing PNG modified'}}
    elseif (Test-Path -LiteralPath $case.Output) {throw 'Unexpected output file'}
    if ($case.Name -eq 'animated' -and (($messages|Out-String) -notmatch 'Animated GIF')) {throw 'Animation was not identified'}
    $results += [pscustomobject]@{Name=$case.Name;Result='PASS'}
}
$results | ConvertTo-Json
