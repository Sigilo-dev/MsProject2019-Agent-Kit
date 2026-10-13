[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing
$InputPath = (Resolve-Path -LiteralPath $InputPath).ProviderPath
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($InputPath) -ine '.gif' -or [IO.Path]::GetExtension($OutputPath) -ine '.png') { throw 'Expected GIF input and PNG output.' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Output exists; choose a new PNG path.' }
# GDI+ may collapse zero-delay duplicate GIF frames. Count image blocks as well.
$gifStream=[IO.File]::OpenRead($InputPath)
$gifReader=[IO.BinaryReader]::new($gifStream)
function Skip-GifBlocks {
    while (($length=$gifReader.ReadByte()) -gt 0) {
        if ($gifStream.Position+$length -gt $gifStream.Length) {throw 'Truncated GIF block.'}
        [void]$gifStream.Seek($length,[IO.SeekOrigin]::Current)
    }
}
try {
    $header=$gifReader.ReadBytes(13)
    if ($header.Length -ne 13 -or [Text.Encoding]::ASCII.GetString($header,0,6) -notin @('GIF87a','GIF89a')) {throw 'Input is not a GIF image.'}
    if ($header[10] -band 128) {[void]$gifStream.Seek(3*[math]::Pow(2,($header[10] -band 7)+1),[IO.SeekOrigin]::Current)}
    $frames=0; $ended=$false
    while ($gifStream.Position -lt $gifStream.Length) {
        $block=$gifReader.ReadByte()
        if ($block -eq 59) {$ended=$true; break}
        if ($block -eq 33) {[void]$gifReader.ReadByte(); Skip-GifBlocks}
        elseif ($block -eq 44) {
            $frames++
            if ($frames -gt 1) {throw 'Animated GIF is not supported; refusing to silently discard frames.'}
            $descriptor=$gifReader.ReadBytes(9)
            if ($descriptor.Length -ne 9) {throw 'Truncated GIF descriptor.'}
            if ($descriptor[8] -band 128) {[void]$gifStream.Seek(3*[math]::Pow(2,($descriptor[8] -band 7)+1),[IO.SeekOrigin]::Current)}
            [void]$gifReader.ReadByte(); Skip-GifBlocks
        } else {throw 'Invalid GIF block.'}
    }
    if (-not $ended -or $frames -ne 1) {throw 'Expected a complete single-frame GIF.'}
} finally {$gifReader.Dispose(); $gifStream.Dispose()}
function Get-PixelHash([Drawing.Image]$Image) {
    $bitmap = [Drawing.Bitmap]::new($Image)
    $data = $null
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        $data = $bitmap.LockBits([Drawing.Rectangle]::new(0,0,$bitmap.Width,$bitmap.Height),[Drawing.Imaging.ImageLockMode]::ReadOnly,[Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $row = New-Object byte[] ($bitmap.Width * 4)
        for ($y=0; $y -lt $bitmap.Height; $y++) {
            [Runtime.InteropServices.Marshal]::Copy([IntPtr]::Add($data.Scan0,$y*$data.Stride),$row,0,$row.Length)
            [void]$hash.TransformBlock($row,0,$row.Length,$row,0)
        }
        [void]$hash.TransformFinalBlock([byte[]]@(),0,0)
        return [BitConverter]::ToString($hash.Hash).Replace('-','')
    } finally {
        if ($data) { $bitmap.UnlockBits($data) }
        $bitmap.Dispose(); $hash.Dispose()
    }
}
$gif = $null; $png = $null
try {
    $gif = [Drawing.Image]::FromFile($InputPath)
    if ($gif.RawFormat.Guid -ne [Drawing.Imaging.ImageFormat]::Gif.Guid) { throw 'Input is not a GIF image.' }
    foreach ($dimension in $gif.FrameDimensionsList) {
        if ($gif.GetFrameCount([Drawing.Imaging.FrameDimension]::new($dimension)) -ne 1) { throw 'Animated GIF is not supported; refusing to silently discard frames.' }
    }
    if ([long]$gif.Width * $gif.Height -gt 100000000) { throw 'Image exceeds the 100-megapixel conversion budget; use a smaller native export.' }
    $pixelHash = Get-PixelHash $gif
    $stream = [IO.File]::Open($OutputPath,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    try { $gif.Save($stream,[Drawing.Imaging.ImageFormat]::Png) } finally { $stream.Dispose() }
    $png = [Drawing.Image]::FromFile($OutputPath)
    if ($png.Width -ne $gif.Width -or $png.Height -ne $gif.Height -or (Get-PixelHash $png) -ne $pixelHash) { throw 'PNG pixel verification failed. Keep the GIF and inspect the output.' }
    [pscustomobject]@{Result='CONVERTED';Input=$InputPath;Output=$OutputPath;Width=$png.Width;Height=$png.Height;PixelIdentical=$true;DecodedArgbSHA256=$pixelHash;Resized=$false} | ConvertTo-Json
} finally { if ($png) {$png.Dispose()}; if ($gif) {$gif.Dispose()} }
