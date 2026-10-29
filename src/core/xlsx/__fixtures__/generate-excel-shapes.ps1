# Generates excel-shapes.xlsx with real Excel: preset shapes, a text box, a connector, a group and
# the form controls (check box, option buttons, combo box, list box, button, spin, scroll bar,
# group box) with linked cells. Run: powershell -File generate-excel-shapes.ps1
$ErrorActionPreference = 'Stop'
while (Get-Process EXCEL -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 30 }
$out = Join-Path $PSScriptRoot 'excel-shapes.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Shapes'
	$ws.Range('A1').Value2 = 'Items'
	$ws.Range('A2').Value2 = 'Red'; $ws.Range('A3').Value2 = 'Green'; $ws.Range('A4').Value2 = 'Blue'
	$r = $ws.Shapes.AddShape(1, 20, 20, 120, 60); $r.Name = 'Box'
	$r.Fill.ForeColor.RGB = 0x0080FF; $r.Line.ForeColor.RGB = 0x000000; $r.Line.Weight = 2
	$r.TextFrame2.TextRange.Text = 'Hello'
	$e = $ws.Shapes.AddShape(9, 160, 20, 100, 60); $e.Name = 'Oval'
	$a = $ws.Shapes.AddShape(33, 280, 20, 100, 60); $a.Name = 'Arrow'
	$t = $ws.Shapes.AddTextbox(1, 20, 100, 140, 40); $t.Name = 'Note'
	$t.TextFrame2.TextRange.Text = 'A text box'
	$c = $ws.Shapes.AddConnector(1, 200, 110, 360, 150); $c.Name = 'Line'
	$g1 = $ws.Shapes.AddShape(1, 20, 170, 50, 30); $g1.Name = 'G1'
	$g2 = $ws.Shapes.AddShape(9, 80, 170, 50, 30); $g2.Name = 'G2'
	$grp = $ws.Shapes.Range(@('G1', 'G2')).Group(); $grp.Name = 'Pair'
	$cb = $ws.CheckBoxes().Add(20, 230, 120, 20); $cb.Caption = 'Enable'; $cb.LinkedCell = '$C$1'; $cb.Value = 1
	$o1 = $ws.OptionButtons().Add(20, 260, 120, 20); $o1.Caption = 'One'; $o1.LinkedCell = '$C$2'
	$o2 = $ws.OptionButtons().Add(20, 280, 120, 20); $o2.Caption = 'Two'; $o2.LinkedCell = '$C$2'
	$dd = $ws.DropDowns().Add(160, 230, 100, 20); $dd.ListFillRange = '$A$2:$A$4'; $dd.LinkedCell = '$C$3'; $dd.Value = 2
	$lb = $ws.ListBoxes().Add(160, 260, 100, 60); $lb.ListFillRange = '$A$2:$A$4'; $lb.LinkedCell = '$C$4'
	$bt = $ws.Buttons().Add(280, 230, 90, 24); $bt.Caption = 'Press me'
	$sp = $ws.Spinners().Add(280, 260, 20, 40); $sp.LinkedCell = '$C$5'; $sp.Min = 0; $sp.Max = 10; $sp.Value = 3
	$sb = $ws.ScrollBars().Add(310, 260, 100, 16); $sb.LinkedCell = '$C$6'; $sb.Min = 0; $sb.Max = 100; $sb.Value = 40
	$gb = $ws.GroupBoxes().Add(400, 230, 120, 80); $gb.Caption = 'Group'
	if (Test-Path $out) { Remove-Item $out }
	$wb.SaveAs($out, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
