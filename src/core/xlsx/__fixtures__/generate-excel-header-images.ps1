# Generates excel-header-images.xlsx with real Excel: header/footer pictures (&G) in every section,
# for the default, first-page and even-page headers, plus a page-number offset (&P+2) and a theme
# colour code. Run: pwsh -File generate-excel-header-images.ps1
$ErrorActionPreference = 'Stop'
while (Get-Process EXCEL -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 30 }
Add-Type -AssemblyName System.Drawing
$png = Join-Path ([IO.Path]::GetTempPath()) 'hf-logo.png'
$bmp = New-Object System.Drawing.Bitmap 40, 20
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255, 200, 40, 40))
$g.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
$out = Join-Path $PSScriptRoot 'excel-header-images.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Print'
	$ws.Range('A1').Value2 = 'Header images'
	$ps = $ws.PageSetup
	$ps.LeftHeaderPicture.Filename = $png
	$ps.LeftHeader = '&G'
	$ps.CenterHeader = 'Page &P+2 of &N'
	$ps.RightHeaderPicture.Filename = $png
	$ps.RightHeader = '&G'
	$ps.CenterFooterPicture.Filename = $png
	$ps.CenterFooter = '&G'
	$ps.DifferentFirstPageHeaderFooter = $true
	$ps.FirstPage.LeftHeader.Picture.Filename = $png
	$ps.FirstPage.LeftHeader.Text = '&G'
	$ps.OddAndEvenPagesHeaderFooter = $true
	$ps.EvenPage.RightFooter.Picture.Filename = $png
	$ps.EvenPage.RightFooter.Text = '&G'
	if (Test-Path $out) { Remove-Item $out }
	$wb.SaveAs($out, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
