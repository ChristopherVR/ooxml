param([Parameter(Mandatory)][string]$InputPath, [Parameter(Mandatory)][string]$OutputDirectory, [string]$FfmpegPath = 'ffmpeg')
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint process before this owned fixture check.' }
$source = (Get-Item -LiteralPath $InputPath).FullName
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$output = (Get-Item -LiteralPath $OutputDirectory).FullName
$video = Join-Path $output 'native-opacity.mp4'
$app = $null
$presentation = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open($source, $false, $false, $false)
    if ($presentation.Slides.Item(1).TimeLine.MainSequence.Count -ne 1) { throw 'Expected one fixed-opacity effect.' }
    if ($presentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text -ne 'Persistent text') { throw 'Unexpected foreground text.' }
    $presentation.SaveCopyAs((Join-Path $output 'native-resaved.pptx'))
    $presentation.CreateVideo($video, $true, 4, 540, 10, 85)
    $deadline = [DateTime]::UtcNow.AddSeconds(60)
    while ($presentation.CreateVideoStatus -ne 3 -and [DateTime]::UtcNow -lt $deadline) {
        if ($presentation.CreateVideoStatus -eq 4) { throw 'Native video export failed.' }
        Start-Sleep -Milliseconds 500
    }
    if ($presentation.CreateVideoStatus -ne 3) { throw 'Native video export timed out.' }
    & $FfmpegPath -hide_banner -loglevel error -i $video -vf fps=5 (Join-Path $output 'frame-%02d.png')
    if ($LASTEXITCODE -ne 0) { throw 'Frame extraction failed.' }
    Add-Type -AssemblyName System.Drawing
    $samples = @()
    foreach ($index in @(1, 4, 6, 10, 15)) {
        $image = [Drawing.Bitmap]::new((Join-Path $output ('frame-' + $index.ToString('00') + '.png')))
        try { $pixel = $image.GetPixel(150, 150) } finally { $image.Dispose() }
        $samples += [pscustomobject]@{ TimestampSeconds=($index-0.5)/5; RGB=@($pixel.R,$pixel.G,$pixel.B) }
    }
    [IO.File]::WriteAllText((Join-Path $output 'frame-samples.json'), ($samples | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    [pscustomobject]@{ NativeVersion=$app.Version; Input=$source; Video=$video; FrameRate=5; Note='Native video has a startup offset; inspect active and completed frames separately from package preservation.' } | ConvertTo-Json
} finally {
    if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
