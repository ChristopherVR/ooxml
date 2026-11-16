<#
.SYNOPSIS
  Capture every built-in SmartArt colour and quick-style definition through PowerPoint COM.

.DESCRIPTION
  Adds one Basic Block List SmartArt per entry of `Application.SmartArtColors` and
  `Application.SmartArtQuickStyles` (documented automation only), applies the entry, nudges the
  size so PowerPoint regenerates the cached drawing, and saves one presentation. PowerPoint writes
  the applied `dgm:colorsDef` / `dgm:styleDef` into each diagram's parts, and the generator
  `gen-smartart-style-definitions.mjs` extracts them by `uniqueId`. The manifest records the
  gallery name, category and description COM reports for each entry, in gallery order.

.PARAMETER OutDir
  Destination of `smartart-definitions.pptx` and `manifest.json`. Defaults to
  `src/core/diagram/__fixtures__/smartart-definitions`.

.NOTES
  Requires a local PowerPoint install (Windows). Deterministic for the same PowerPoint build.
#>
param([string]$OutDir = "$PSScriptRoot\..\..\diagram\__fixtures__\smartart-definitions")
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$app = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
  # The SmartArt galleries stay empty until PowerPoint has a document window.
  $app.Visible = -1
  $pres = $app.Presentations.Add()
  $pres.Slides.Add(1, 12) | Out-Null
  Start-Sleep -Milliseconds 800
  $pres.Slides.Item(1).Delete()
  $layout = $null
  for ($i = 1; $i -le $app.SmartArtLayouts.Count; $i++) {
    if ([string]$app.SmartArtLayouts.Item($i).Id -eq 'urn:microsoft.com/office/officeart/2005/8/layout/default') {
      $layout = $app.SmartArtLayouts.Item($i)
    }
  }
  if (-not $layout) { throw 'Basic Block List is not available.' }
  $entries = @()
  $slideIndex = 1
  foreach ($kind in @('colors', 'style')) {
    # Assigned directly: an if-expression would enumerate the COM collection.
    $collection = $app.SmartArtQuickStyles
    if ($kind -eq 'colors') { $collection = $app.SmartArtColors }
    for ($i = 1; $i -le $collection.Count; $i++) {
      $item = $collection.Item($i)
      $slide = $pres.Slides.Add($slideIndex, 12)
      $shape = $slide.Shapes.AddSmartArt($layout, 60, 60, 600, 360)
      $art = $shape.SmartArt
      if ($kind -eq 'colors') { $art.Color = $item } else { $art.QuickStyle = $item }
      $w = [single]$shape.Width; $h = [single]$shape.Height
      $shape.Width = [single]($w + 5); $shape.Height = [single]($h + 5)
      $shape.Width = $w; $shape.Height = $h
      $entries += [ordered]@{
        slide = $slideIndex
        kind = $kind
        order = $i
        id = [string]$item.Id
        name = [string]$item.Name
        category = [string]$item.Category
        description = [string]$item.Description
      }
      $slideIndex++
    }
  }
  $pres.SaveAs((Join-Path $OutDir 'smartart-definitions.pptx'), 24)
  [ordered]@{ powerpoint = [string]$app.Version; build = [string]$app.Build; entries = $entries } |
    ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $OutDir 'manifest.json')
  Write-Output "Captured $($entries.Count) definitions with PowerPoint $($app.Version)."
} finally {
  if ($pres) { $pres.Close() }
  $app.Quit()
}
# Keep the local user's name out of the committed fixture: the author fields become `Author`.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open((Join-Path $OutDir 'smartart-definitions.pptx'), 'Update')
try {
  $entry = $zip.GetEntry('docProps/core.xml')
  $reader = New-Object System.IO.StreamReader($entry.Open())
  $xml = $reader.ReadToEnd(); $reader.Close()
  $xml = $xml -replace '<dc:creator>[^<]*</dc:creator>', '<dc:creator>Author</dc:creator>'
  $xml = $xml -replace '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>', '<cp:lastModifiedBy>Author</cp:lastModifiedBy>'
  $entry.Delete()
  $writer = New-Object System.IO.StreamWriter($zip.CreateEntry('docProps/core.xml').Open())
  $writer.Write($xml); $writer.Close()
} finally { $zip.Dispose() }
