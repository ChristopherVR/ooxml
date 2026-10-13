param([Parameter(Mandatory)][string]$SourcePath, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint process before this owned fixture check.' }
$source = (Get-Item -LiteralPath $SourcePath).FullName
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$output = (Get-Item -LiteralPath $OutputDirectory).FullName
$wave = Join-Path $output 'fixture-silence.wav'
$stream = [IO.MemoryStream]::new()
$writer = [IO.BinaryWriter]::new($stream)
try {
    $writer.Write([Text.Encoding]::ASCII.GetBytes('RIFF'))
    $writer.Write([int]16036)
    $writer.Write([Text.Encoding]::ASCII.GetBytes('WAVEfmt '))
    $writer.Write([int]16)
    $writer.Write([int16]1)
    $writer.Write([int16]1)
    $writer.Write([int]8000)
    $writer.Write([int]16000)
    $writer.Write([int16]2)
    $writer.Write([int16]16)
    $writer.Write([Text.Encoding]::ASCII.GetBytes('data'))
    $writer.Write([int]16000)
    $writer.Write([byte[]]::new(16000))
    [IO.File]::WriteAllBytes($wave, $stream.ToArray())
} finally { $writer.Dispose(); $stream.Dispose() }
$app = $null
$presentation = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open($source, $false, $false, $false)
    $presentation.Slides.Item(1).Shapes.Item(1).AnimationSettings.SoundEffect.ImportFromFile($wave)
    $presentation.SaveCopyAs((Join-Path $output 'native-after-effect-with-sound.pptx'))
} finally {
    if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
