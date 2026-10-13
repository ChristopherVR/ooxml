param([Parameter(Mandatory)][string]$InputPath, [Parameter(Mandatory)][string]$OutputDirectory, [switch]$WholeOpacity)
$ErrorActionPreference = 'Stop'
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint process before this owned fixture check.' }
$source = (Get-Item -LiteralPath $InputPath).FullName
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$output = (Get-Item -LiteralPath $OutputDirectory).FullName
$resaved = Join-Path $output 'background-native-resaved.pptx'
$app = $null
$presentation = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $presentation = $app.Presentations.Open($source, $false, $false, $false)
    $effects = $presentation.Slides.Item(1).TimeLine.MainSequence.Count
    $text = $presentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text
    $expectedEffects = if ($WholeOpacity) { 2 } else { 1 }
    if ($effects -ne $expectedEffects -or $text -ne 'Persistent text') { throw 'Unexpected synthetic source content or effects.' }
    $presentation.Slides.Item(1).Export((Join-Path $output 'background-native.png'), 'PNG', 1280, 720)
    $presentation.SaveCopyAs($resaved)
    $presentation.Close()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
    $presentation = $null
    $presentation = $app.Presentations.Open($resaved, $false, $false, $false)
    $reopenedEffects = $presentation.Slides.Item(1).TimeLine.MainSequence.Count
    $reopenedText = $presentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text
    if ($reopenedEffects -ne $effects -or $reopenedText -ne $text) { throw 'Native resave changed the synthetic effect or text.' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($resaved)
    try {
        $reader = [IO.StreamReader]::new($archive.GetEntry('ppt/slides/slide1.xml').Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $targets = $xml.SelectNodes("//*[local-name()='spTgt']/*[local-name()='bg']").Count
        $wholeTargets = $xml.SelectNodes("//*[local-name()='spTgt'][not(*[local-name()='bg'])]").Count
        if ($targets -lt 1 -or (!$WholeOpacity -and $wholeTargets -ne 0)) { throw 'Native resave did not retain background-only shape targets.' }
        if ($WholeOpacity -and ($wholeTargets -lt 1 -or $xml.SelectNodes("//*[local-name()='attrName' and text()='style.opacity']").Count -lt 1)) { throw 'Native resave lost the concurrent whole-shape opacity behavior.' }
    } finally { $archive.Dispose() }
    [pscustomobject]@{ NativeVersion=$app.Version; SourceEffects=$effects; ReopenedEffects=$reopenedEffects; BackgroundTargets=$targets; Text=$reopenedText } | ConvertTo-Json
} finally {
    if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) }
    if ($app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
