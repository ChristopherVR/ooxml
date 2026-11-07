<#
.SYNOPSIS
  Unhinted glyph advances per 1000 em for printable ASCII (32..126), regular and bold, measured
  through real PowerPoint COM, printed as JSON for `design-advances.generated.ts`.

.DESCRIPTION
  Same method as make-font-advance-table.ps1: one non-wrapping zero-inset text box at 100pt,
  advance(g) = BoundWidth("H" + g + "H") - BoundWidth("HH"). Used for the theme fonts Visio
  does not have (Aptos), so the DOM text layer can give a stand-in font the deck font's widths.
  Requires a local PowerPoint install (Windows).
#>
param([string[]]$Fonts = @('Aptos', 'Aptos Display'))
$ErrorActionPreference = 'Stop'
$ppt = New-Object -ComObject PowerPoint.Application
$deck = $ppt.Presentations.Add($false)
$slide = $deck.Slides.Add(1, 12)
$box = $slide.Shapes.AddTextbox(1, 0, 0, 5000, 200)
$box.TextFrame.WordWrap = 0
$box.TextFrame.MarginLeft = 0; $box.TextFrame.MarginRight = 0
$result = [ordered]@{}
foreach ($font in $Fonts) {
  $faces = [ordered]@{}
  foreach ($bold in @($false, $true)) {
    $range = $box.TextFrame.TextRange
    $range.Font.Name = $font; $range.Font.Size = 100; $range.Font.Bold = [int]$bold * -1
    $range.Text = 'HH'
    $base = $range.BoundWidth
    $values = for ($code = 32; $code -le 126; $code++) {
      $range.Text = 'H' + [char]$code + 'H'
      [math]::Round(($range.BoundWidth - $base) * 10, 2)
    }
    $faces[$(if ($bold) { 'bold' } else { 'regular' })] = ($values -join ' ')
  }
  $result[$font.ToLowerInvariant()] = $faces
}
$deck.Close(); $ppt.Quit()
$result | ConvertTo-Json -Depth 4
