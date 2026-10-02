# Generates the Excel-authored fixtures used by the xlsx read/write tests.
# Requires desktop Excel (COM automation). Run from any folder:
#   pwsh -File src/xlsx/__fixtures__/generate-excel.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$png = Join-Path $env:TEMP 'xlsx-fixture-image.png'
# A small solid PNG drawn with System.Drawing (Windows only, like Excel itself).
Add-Type -AssemblyName System.Drawing
$bitmap = New-Object System.Drawing.Bitmap 4, 4
for ($x = 0; $x -lt 4; $x++) { for ($y = 0; $y -lt 4; $y++) { $bitmap.SetPixel($x, $y, [System.Drawing.Color]::SteelBlue) } }
$bitmap.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

# Excel stamps the signed-in account and the save folder into the package; replace them so the
# committed fixtures carry no personal data.
function Remove-PersonalData([string]$path) {
	Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
	$zip = [IO.Compression.ZipFile]::Open($path, 'Update')
	try {
		$edits = @(
			@{ Part = 'xl/persons/person.xml'; Find = 'displayName="[^"]*"'; With = 'displayName="Fixture Person"' }
			@{ Part = 'xl/persons/person.xml'; Find = 'userId="[^"]*"'; With = 'userId="fixture"' }
			@{ Part = 'xl/persons/person.xml'; Find = 'providerId="[^"]*"'; With = 'providerId="None"' }
			@{ Part = 'docProps/core.xml'; Find = '<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>'; With = '<cp:lastModifiedBy>Fixture Author</cp:lastModifiedBy>' }
			@{ Part = 'docProps/core.xml'; Find = '<dc:creator>[^<]*</dc:creator>'; With = '<dc:creator>Fixture Author</dc:creator>' }
			@{ Part = 'xl/workbook.xml'; Find = '<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath[^>]*/></mc:Choice></mc:AlternateContent>'; With = '' }
		)
		foreach ($name in ($edits | ForEach-Object { $_.Part } | Select-Object -Unique)) {
			$entry = $zip.GetEntry($name)
			if (-not $entry) { continue }
			$reader = New-Object IO.StreamReader($entry.Open())
			$text = $reader.ReadToEnd()
			$reader.Close()
			foreach ($edit in ($edits | Where-Object { $_.Part -eq $name })) { $text = [regex]::Replace($text, $edit.Find, $edit.With) }
			$entry.Delete()
			$writer = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
			$writer.Write($text)
			$writer.Close()
		}
	} finally { $zip.Dispose() }
}

$excel = New-Object -ComObject Excel.Application
# Excel.UserName is the user's Office-wide name (saved in the registry): stamp the
# fixtures with a neutral name, then put the user's own name back in the finally block.
$previousUserName = $excel.UserName
$excel.UserName = 'Fixture Author'
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	# ---- excel-features.xlsx ------------------------------------------------------------
	$wb = $excel.Workbooks.Add()
	while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Main'
	$ws.Range('A1').Value2 = 'Item'
	$ws.Range('B1').Value2 = 'Qty'
	$ws.Range('C1').Value2 = 'Price'
	$ws.Range('D1').Value2 = 'Total'
	$items = @('Apple', 'Banana', 'Cherry', 'Date', 'Elder')
	for ($i = 0; $i -lt 5; $i++) {
		$r = $i + 2
		$ws.Range("A$r").Value2 = $items[$i]
		$ws.Range("B$r").Value2 = [double](($i + 1) * 3)
		$ws.Range("C$r").Value2 = [double](1.25 * ($i + 1))
	}
	# Filling one formula down a range makes Excel store a shared formula.
	$ws.Range('D2:D6').Formula = '=B2*C2'
	$ws.Range('E2:E6').Formula = '=D2/SUM($D$2:$D$6)'
	$ws.Range('E2:E6').NumberFormat = '0.0%'
	$ws.Range('B8').Formula = '=SUM(B2:B6)'
	$ws.Range('C8').FormulaArray = '=MAX(B2:B6*C2:C6)'
	$ws.Range('D8').Formula = '=IF(B8>10,"big","small")'
	$ws.Range('A10').Value2 = 'Rich text'
	$ws.Range('A10').Characters(1, 4).Font.Bold = $true
	$ws.Range('A10').Characters(6, 4).Font.Color = 255
	$ws.Range('A11').Value2 = "Line 1`r`nLine 2"
	$ws.Range('A12').Value2 = [double]45000.5
	$ws.Range('A12').NumberFormat = 'dd/mm/yyyy hh:mm'
	$ws.Range('B12').Value2 = -1234.5
	$ws.Range('B12').NumberFormat = '#,##0.00_);[Red](#,##0.00)'
	$ws.Range('C12').Value2 = $true
	$ws.Range('D12').Formula = '=1/0'
	$ws.Range('A1:D1').Font.Bold = $true
	$ws.Range('A1:D1').Interior.Color = 0xC07000
	$ws.Range('A1:D1').Font.ThemeColor = 1
	$ws.Range('A1:D1').Borders.Item(9).LineStyle = 1
	$ws.Range('A1:D1').Borders.Item(9).Weight = -4138
	$ws.Range('B2:B6').Interior.ThemeColor = 6
	$ws.Range('B2:B6').Interior.TintAndShade = 0.6
	$ws.Range('A14:C15').Merge()
	$ws.Range('A14').Value2 = 'Merged block'
	$ws.Range('A14').HorizontalAlignment = -4108
	$ws.Columns.Item(1).ColumnWidth = 16
	$ws.Columns.Item('G:H').ColumnWidth = 4
	$ws.Rows.Item(11).RowHeight = 30
	$ws.Rows.Item(13).Hidden = $true
	$ws.Activate()
	$excel.ActiveWindow.FreezePanes = $false
	$ws.Range('B2').Select()
	$excel.ActiveWindow.FreezePanes = $true
	$excel.ActiveWindow.Zoom = 90
	$ws.Tab.Color = 0x00B050
	[void]$ws.Range('B2:B6').FormatConditions.Add(1, 5, '=6')
	$ws.Range('B2:B6').FormatConditions.Item(1).Interior.Color = 0x9CC7FF
	[void]$ws.Range('C2:C6').FormatConditions.AddDatabar()
	[void]$ws.Range('D2:D6').FormatConditions.AddColorScale(3)
	[void]$ws.Range('E2:E6').FormatConditions.AddIconSetCondition()
	$dv = $ws.Range('F2:F6').Validation
	$dv.Add(3, 1, 1, 'Yes,No')
	$dv.InputTitle = 'Answer'
	$dv.InputMessage = 'Yes or No'
	$dv2 = $ws.Range('G2:G6').Validation
	$dv2.Add(1, 1, 1, '1', '10')
	[void]$ws.Range('A2').AddComment('Legacy note text')
	try { [void]$ws.Range('A3').AddCommentThreaded('Threaded root'); [void]$ws.Range('A3').CommentThreaded.AddReply('A reply') } catch { Write-Host 'threaded comments unavailable' }
	[void]$ws.Hyperlinks.Add($ws.Range('H2'), 'https://example.org/', '', 'Example tip', 'Example')
	[void]$ws.Hyperlinks.Add($ws.Range('H3'), '', "'Second'!A1", '', 'Go to Second')
	$ws.PageSetup.Orientation = 2
	$ws.PageSetup.CenterHeader = 'Header text'
	$ws.PageSetup.PrintArea = '$A$1:$E$12'
	[void]$wb.Names.Add('Rate', '=Main!$C$2')

	$s2 = $wb.Worksheets.Item(2)
	$s2.Name = 'Second'
	$s2.Range('A1').Value2 = 'Region'
	$s2.Range('B1').Value2 = 'Sales'
	$s2.Range('C1').Value2 = 'Cost'
	$regions = @('North', 'South', 'East', 'West')
	for ($i = 0; $i -lt 4; $i++) {
		$r = $i + 2
		$s2.Range("A$r").Value2 = $regions[$i]
		$s2.Range("B$r").Value2 = [double](100 + $i * 25)
		$s2.Range("C$r").Value2 = [double](60 + $i * 10)
	}
	$lo = $s2.ListObjects.Add(1, $s2.Range('A1:C5'), $null, 1)
	$lo.Name = 'SalesTable'
	$lo.TableStyle = 'TableStyleMedium2'
	$lo.ShowTotals = $true
	$chartShape = $s2.Shapes.AddChart2(201, 51, 300, 10, 360, 220)
	$chartShape.Chart.SetSourceData($s2.Range('A1:C5'))
	$chartShape.Chart.HasTitle = $true
	$chartShape.Chart.ChartTitle.Text = 'Sales by region'
	[void]$s2.Shapes.AddPicture($png, 0, -1, 10, 200, 40, 40)
	[void]$s2.Shapes.AddShape(1, 100, 200, 80, 40)

	$s3 = $wb.Worksheets.Item(3)
	$s3.Name = 'Secret'
	$s3.Range('A1').Value2 = 'hidden value'
	$s3.Visible = [int]0
	$main = $wb.Worksheets.Item(1)
	$main.Activate()
	try { $wb.Title = 'Excel fixture'; $wb.Author = 'Excel generator' } catch { Write-Host 'document properties not set' }
	$wb.SaveAs((Join-Path $here 'excel-features.xlsx'), 51)
	$wb.Close($false)

	# ---- excel-1904.xlsx ----------------------------------------------------------------
	$wb = $excel.Workbooks.Add()
	$wb.Date1904 = $true
	$ws = $wb.Worksheets.Item(1)
	$ws.Range('A1').Value2 = [double]0
	$ws.Range('A1').NumberFormat = 'yyyy-mm-dd'
	$ws.Range('A2').Formula = '=DATE(2020,6,1)'
	$ws.Range('A2').NumberFormat = 'yyyy-mm-dd'
	$wb.SaveAs((Join-Path $here 'excel-1904.xlsx'), 51)
	$wb.Close($false)
} finally {
	$excel.UserName = $previousUserName
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
	Remove-Item $png -ErrorAction SilentlyContinue
}
foreach ($file in Get-ChildItem $here -Filter 'excel-*.xlsx') { Remove-PersonalData $file.FullName }
Get-ChildItem $here -Filter 'excel-*.xlsx' | Select-Object Name, Length
