# Capture native PowerPoint shapes named by a generated visual manifest.
# Export a full slide at the requested density, then crop the declared shape frame in compare.ts.
# Shape.Export adds ink bounds and uses point-based scaling, so it is not a pixel-exact viewport.
param([Parameter(Mandatory=$true)][string]$Manifest)
$ErrorActionPreference = 'Stop'
$taskManifestPath = (Resolve-Path -LiteralPath $Manifest).Path
$taskRoot = Split-Path -Parent $taskManifestPath
$taskSpec = Get-Content -LiteralPath $taskManifestPath -Raw | ConvertFrom-Json
foreach ($taskEntry in $taskSpec.cases) {
 if (-not $taskEntry.presentation) { $taskEntry | Add-Member -NotePropertyName presentation -NotePropertyValue $taskSpec.presentation -Force }
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close the existing PowerPoint session before this native acceptance run.' }
$taskResults = @()
$taskApp = New-Object -ComObject PowerPoint.Application
$taskApp.DisplayAlerts = 1
try {
 foreach ($taskGroup in ($taskSpec.cases | Group-Object presentation)) {
  $taskSourcePath = if ([IO.Path]::IsPathRooted($taskGroup.Name)) { $taskGroup.Name } else { Join-Path $taskRoot $taskGroup.Name }
  $taskSource = (Resolve-Path -LiteralPath $taskSourcePath).Path
  $taskDeck = $null
  try {
   $taskDeck = $taskApp.Presentations.Open($taskSource, -1, 0, 0)
   foreach ($taskCase in $taskGroup.Group) {
    if ($taskCase.id -notmatch '^[a-z0-9][a-z0-9-]*$') { throw "Invalid case id: $($taskCase.id)" }
    if ($taskCase.width -lt 1 -or $taskCase.width -gt 4096 -or $taskCase.height -lt 1 -or $taskCase.height -gt 4096) { throw 'Capture dimensions must be between 1 and 4096.' }
    $taskShape = $taskDeck.Slides.Item([int]$taskCase.slide).Shapes.Item([int]$taskCase.shape)
    $taskWidth = [int][Math]::Round($taskDeck.PageSetup.SlideWidth * $taskCase.width / $taskShape.Width)
    $taskHeight = [int][Math]::Round($taskDeck.PageSetup.SlideHeight * $taskCase.height / $taskShape.Height)
    $taskOutput = Join-Path $taskRoot "$($taskCase.id).native-slide.png"
    $taskDeck.Slides.Item([int]$taskCase.slide).Export($taskOutput, 'PNG', $taskWidth, $taskHeight)
    $taskResults += [ordered]@{
     id = $taskCase.id; nativeSlidePng = [IO.Path]::GetFileName($taskOutput)
     crop = @{ x = [int][Math]::Round($taskShape.Left * $taskWidth / $taskDeck.PageSetup.SlideWidth); y = [int][Math]::Round($taskShape.Top * $taskHeight / $taskDeck.PageSetup.SlideHeight); width = $taskCase.width; height = $taskCase.height }
     sourceSha256 = (Get-FileHash -LiteralPath $taskSource -Algorithm SHA256).Hash
     widthPt = $taskShape.Width; heightPt = $taskShape.Height; rotation = $taskShape.Rotation
     slide = $taskCase.slide; shape = $taskCase.shape
    }
   }
  } finally { if ($null -ne $taskDeck) { $taskDeck.Saved = -1; $taskDeck.Close() } }
 }
 $taskExe = Join-Path $taskApp.Path 'POWERPNT.EXE'
 $taskEvidence = [ordered]@{
  capturedAt = [DateTime]::UtcNow.ToString('o'); officeVersion = $taskApp.Version
  officeFileVersion = (Get-Item -LiteralPath $taskExe).VersionInfo.FileVersion
  capture = 'PowerPoint Slide.Export PNG at requested density, declared shape-frame crop without image registration'
  cases = $taskResults
 }
 $taskEvidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $taskRoot 'native.json') -Encoding utf8
 $taskEvidence | ConvertTo-Json -Depth 8
} finally { $taskApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($taskApp) }
