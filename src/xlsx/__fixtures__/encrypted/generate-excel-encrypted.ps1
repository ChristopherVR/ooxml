# Generates excel-encrypted.xlsx: a workbook Excel saved with the password to open 'open sesame'
# (agile encryption), used by the encrypted-package load tests. Requires desktop Excel (COM):
#   pwsh -File src/xlsx/__fixtures__/encrypted/generate-excel-encrypted.ps1
# The package is encrypted, so it cannot be scrubbed afterwards like the other fixtures:
# RemovePersonalInformation makes Excel drop the author and last-modified-by names on save.
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$path = Join-Path $here 'excel-encrypted.xlsx'
if (Test-Path $path) { Remove-Item -Force $path }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = 'Secret'
	$ws.Range('A1').Value2 = 'Item'
	$ws.Range('B1').Value2 = 'Qty'
	$ws.Range('A2').Value2 = 'Apple'
	$ws.Range('B2').Value2 = 3
	$ws.Range('A3').Value2 = 'Banana'
	$ws.Range('B3').Value2 = 4
	$ws.Range('B4').Formula = '=SUM(B2:B3)'
	$wb.RemovePersonalInformation = $true
	$wb.Password = 'open sesame'
	# 51 = xlOpenXMLWorkbook (.xlsx)
	$wb.SaveAs($path, 51)
	$wb.Close($false)
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
Write-Output "wrote $path"
