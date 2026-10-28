param(
 [Parameter(Mandatory=$true)][string]$InputPresentation,
 [Parameter(Mandatory=$true)][string]$OutputPresentation,
 [int]$Slide=1,
 [int]$Shape=1,
 [int]$EffectCode=21
)
# Reproduce the committed native-artistic-fill.pptx from a generated pictures.pptx.
# Installed Office produced a14:artisticPencilSketch for Fill.PictureEffects.Insert(21).
# Record the emitted XML rather than inferring a preset name from disputed enum tables.
$ErrorActionPreference='Stop'
if(Get-Process POWERPNT -ErrorAction SilentlyContinue){throw 'Close PowerPoint before running the native authoring probe.'}
$inputPath=(Resolve-Path -LiteralPath $InputPresentation).Path
$outputPath=[System.IO.Path]::GetFullPath($OutputPresentation)
$app=New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts=1
$deck=$null
try {
 $deck=$app.Presentations.Open($inputPath,0,0,0)
 $effect=$deck.Slides.Item($Slide).Shapes.Item($Shape).Fill.PictureEffects.Insert($EffectCode)
 $deck.SaveAs($outputPath,24)
} finally {
 if($deck){$deck.Saved=-1;$deck.Close()}
 $app.Quit()
}
