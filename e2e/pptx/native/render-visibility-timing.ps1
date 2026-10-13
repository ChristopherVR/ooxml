param([Parameter(Mandatory)][string]$Directory, [string]$FfmpegPath = 'ffmpeg')
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint process before this owned fixture check.' }
$folder = (Get-Item -LiteralPath $Directory).FullName
$app = $null
$presentation = $null
$evidence = @()
try {
    $app = New-Object -ComObject PowerPoint.Application
    foreach ($name in @('zero','genuine','long')) {
        $source = (Get-Item -LiteralPath (Join-Path $folder ($name+'.pptx'))).FullName
        $presentation = $app.Presentations.Open($source, $false, $false, $false)
        $sequence = $presentation.Slides.Item(1).TimeLine.MainSequence
        if ($sequence.Count -ne 2) { throw 'Unexpected native effect count.' }
        $effect = $sequence.Item(1)
        if ($effect.Shape.Id -ne 2 -or $effect.EffectInformation.AfterEffect -ne 0) { throw 'Unexpected native visibility target or after-state.' }
        $nativeDuration = $effect.Timing.Duration
        $presentation.SaveCopyAs((Join-Path $folder ($name+'-native.pptx')))
        $video = Join-Path $folder ($name+'.mp4')
        $presentation.CreateVideo($video, $true, 4, 540, 20, 85)
        $deadline = [DateTime]::UtcNow.AddSeconds(90)
        while ($presentation.CreateVideoStatus -ne 3 -and [DateTime]::UtcNow -lt $deadline) {
            if ($presentation.CreateVideoStatus -eq 4) { throw 'Native video export failed.' }
            Start-Sleep -Milliseconds 250
        }
        if ($presentation.CreateVideoStatus -ne 3) { throw 'Native video export timed out.' }
        $frames = Join-Path $folder ($name+'-frames')
        [IO.Directory]::CreateDirectory($frames) | Out-Null
        & $FfmpegPath -hide_banner -loglevel error -i $video -vf fps=20 (Join-Path $frames 'frame-%04d.png')
        if ($LASTEXITCODE -ne 0) { throw 'Native frame extraction failed.' }
        Add-Type -AssemblyName System.Drawing
        $samples = @()
        foreach ($frame in Get-ChildItem -LiteralPath $frames -Filter '*.png' | Sort-Object Name) {
            $image = [Drawing.Bitmap]::new($frame.FullName)
            try { $pixel = $image.GetPixel(80, 80); $marker = $image.GetPixel(80, 280) } finally { $image.Dispose() }
            $index = [int]($frame.BaseName -replace '\D','')
            $samples += @{TimestampSeconds=($index-0.5)/20;RGB=@($pixel.R,$pixel.G,$pixel.B);MarkerRGB=@($marker.R,$marker.G,$marker.B)}
        }
        $evidence += @{Case=$name;NativeDurationSeconds=$nativeDuration;NativeVersion=$app.Version;Samples=$samples;Note='The concurrent zero-delay marker establishes the actual effect start independently of movie startup and slide pacing.'}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($effect)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sequence)
        $presentation.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
        $presentation = $null
        $presentation = $app.Presentations.Open((Join-Path $folder ($name+'-native.pptx')), $false, $false, $false)
        $reopened = $presentation.Slides.Item(1).TimeLine.MainSequence
        if ($reopened.Count -ne 2 -or [math]::Abs($reopened.Item(1).Timing.Duration - $nativeDuration) -gt 0.00001) { throw 'Native reopened effect count or timing changed.' }
        $evidence[-1]['NativeReopened'] = $true
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reopened)
        $presentation.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
        $presentation = $null
    }
    [IO.File]::WriteAllText((Join-Path $folder 'visibility-native-evidence.json'),($evidence|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
} finally {
    if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
