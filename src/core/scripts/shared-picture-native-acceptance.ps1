# Run after fixture generation, while no other acceptance agent owns Office COM.
# Native property/normal-open evidence only: no pixel-fidelity or repair-detection claim.
param([Parameter(Mandatory=$true)][string]$InputDirectory, [string]$EvidencePath)
$ErrorActionPreference = 'Stop'
$taskInput = (Resolve-Path -LiteralPath $InputDirectory).Path
$taskEvidence = [ordered]@{ timestamp = [DateTime]::UtcNow.ToString('o'); limits = 'Normal open and native properties only; alpha pixel output and repair detection are not measured.' }
$taskWord = New-Object -ComObject Word.Application
try {
 $taskWord.Visible = $false; $taskWord.DisplayAlerts = 0
 $taskDocument = $taskWord.Documents.Open((Join-Path $taskInput 'custom-shape.docx'), $false, $true)
 try {
  $taskShape = $taskDocument.Shapes.Item(1)
  $taskFirst = $taskShape.Nodes.Item(1).Points; $taskThird = $taskShape.Nodes.Item(3).Points
  $taskFraction = ($taskFirst.GetValue(1,1) - $taskThird.GetValue(1,1)) / $taskShape.Width
  if ([Math]::Abs($taskFraction - 0.125) -gt 0.001) { throw "Unexpected guide coordinate fraction: $taskFraction" }
  $taskEvidence.word = @{ version = $taskWord.Version; build = $taskWord.Build; nodes = $taskShape.Nodes.Count; widthPt = $taskShape.Width; guideFraction = $taskFraction }
 } finally { $taskDocument.Close(0) }
} finally { $taskWord.Quit() }
$taskExcel = New-Object -ComObject Excel.Application
try {
 $taskExcel.Visible = $false; $taskExcel.DisplayAlerts = $false
 $taskBook = $taskExcel.Workbooks.Open((Join-Path $taskInput 'picture-effects.xlsx'), 0, $true)
 try {
  $taskPicture = $null
  foreach ($taskSheet in $taskBook.Worksheets) { foreach ($taskCandidate in $taskSheet.Shapes) { if ($taskCandidate.Type -eq 13) { $taskPicture = $taskCandidate; break } } if ($null -ne $taskPicture) { break } }
  if ($null -eq $taskPicture) { throw 'The workbook has no native picture' }
  $taskFormat = $taskPicture.PictureFormat
  if ([Math]::Abs($taskPicture.Width - 75) -gt 0.01 -or [Math]::Abs($taskPicture.Height - 37.5) -gt 0.01) { throw "Native resized picture dimensions differ: $($taskPicture.Width)x$($taskPicture.Height)" }
  if ([Math]::Abs($taskFormat.Brightness - 0.6) -gt 0.0001 -or [Math]::Abs($taskFormat.Contrast - 0.45) -gt 0.0001 -or $taskFormat.ColorType -ne 2) { throw 'Native picture correction properties differ' }
  $taskCropRatio = $taskFormat.CropLeft / $taskFormat.CropRight
  if ([Math]::Abs($taskCropRatio - 2.5) -gt 0.001) { throw "Unexpected native crop ratio: $taskCropRatio" }
  $taskEvidence.excel = @{ version = $taskExcel.Version; build = $taskExcel.Build; brightness = $taskFormat.Brightness; contrast = $taskFormat.Contrast; colorType = $taskFormat.ColorType; widthPt = $taskPicture.Width; heightPt = $taskPicture.Height; cropLeftPt = $taskFormat.CropLeft; cropRightPt = $taskFormat.CropRight }
 } finally { $taskBook.Close($false) }
} finally { $taskExcel.Quit() }
$taskJson = $taskEvidence | ConvertTo-Json -Depth 5
if ($EvidencePath) { Set-Content -LiteralPath $EvidencePath -Value $taskJson -Encoding utf8 }
$taskJson
