# Author drawing shapes (preset geometry, fill, outline, text, wrapping) in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-shapes-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$document = $null
try {
    $document = $word.Documents.Add()
    $document.Content.Text = "First paragraph holds the shapes.`rSecond paragraph of text."
    $anchor = $document.Paragraphs.Item(1).Range
    # msoShapeRectangle 1, msoShapeRoundedRectangle 5, msoShapeOval 9, msoShapeRightArrow 33, msoShapeFlowchartProcess 61
    $rect = $document.Shapes.AddShape(1, 72, 90, 140, 70, $anchor)
    $rect.Fill.ForeColor.RGB = 0x0000FF  # BGR: red
    $rect.Line.ForeColor.RGB = 0x00FF00
    $rect.Line.Weight = 3
    $rect.TextFrame.TextRange.Text = 'Red box'
    $rect.WrapFormat.Type = 3   # wdWrapFront
    $arrow = $document.Shapes.AddShape(33, 260, 90, 120, 60, $anchor)
    $arrow.WrapFormat.Type = 0  # wdWrapSquare
    $oval = $document.Shapes.AddShape(9, 72, 200, 100, 100, $anchor)
    $oval.Fill.Visible = 0
    $oval.WrapFormat.Type = 3
    $line = $document.Shapes.AddLine(200, 220, 380, 280, $anchor)
    $line.Line.Weight = 2
    $line.WrapFormat.Type = 3
    $document.SaveAs2((Join-Path $destination 'shapes.docx'), 12)
    $report = [ordered]@{ applicationVersion = [string]$word.Version; shapes = $document.Shapes.Count }
    $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}
