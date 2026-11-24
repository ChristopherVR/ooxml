# Makes source/dessins.docx with an owned Microsoft Word 16 through its documented COM object
# model (Shapes.AddShape, AddLine, BuildFreeform, AddChart2): a rectangle, an ellipse, a line, a
# freeform with a curve and a clustered bar chart under two lines of text. Export it to PDF with
# `export-word.ps1 -Names dessins`. Run on Windows with Word installed, with PowerShell 7.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	$doc = $word.Documents.Add()
	$doc.Content.Text = "Dessins vectoriels`rUn rectangle, une ellipse, une ligne, une forme libre et un graphique en barres suivent.`r"
	$shapes = $doc.Shapes
	# 1 = msoShapeRectangle, 9 = msoShapeOval; colours are BGR.
	$rectangle = $shapes.AddShape(1, 72, 150, 120, 60)
	$rectangle.Fill.ForeColor.RGB = 0xCC6633
	$rectangle.Line.ForeColor.RGB = 0x000000
	$oval = $shapes.AddShape(9, 230, 150, 90, 60)
	$oval.Fill.ForeColor.RGB = 0x33AA33
	$oval.Line.Visible = 0
	$line = $shapes.AddLine(360, 150, 480, 210)
	$line.Line.Weight = 2
	$line.Line.ForeColor.RGB = 0x0000FF
	# 0 = msoEditingAuto, msoSegmentLine; 1 = msoSegmentCurve.
	$builder = $shapes.BuildFreeform(0, 72, 260)
	$builder.AddNodes(0, 0, 140, 300)
	$builder.AddNodes(1, 0, 200, 240, 260, 320, 300, 260)
	$builder.AddNodes(0, 0, 72, 260)
	$freeform = $builder.ConvertToShape()
	$freeform.Fill.ForeColor.RGB = 0x9933CC
	# 57 = xlBarClustered, with Word's default sample data.
	$null = $shapes.AddChart2(-1, 57, 72, 360, 360, 220)
	$doc.SaveAs2((Join-Path $here 'source/dessins.docx'), 16)
	$doc.Close(0)
} finally { $word.Quit() }
