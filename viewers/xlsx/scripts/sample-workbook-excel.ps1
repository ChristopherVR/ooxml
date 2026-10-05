# Lets real Excel recalculate and re-save the demo sample workbook, and writes the legacy
# Excel 97-2003 fixture the package smoke test and the browser specs load.
#
#   python scripts/sample-workbook.py
#   pwsh scripts/sample-workbook-excel.ps1
#
# Manual (needs Excel on Windows); the outputs are committed, CI never runs this.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$sample = Join-Path $root 'demos\demo-vanilla\public\sample.xlsx'
$legacy = Join-Path $root 'tests\support\legacy-97.xls'
$xlOpenXMLWorkbook = 51
$xlExcel8 = 56

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Excel.UserName is the user's Office-wide name (saved in the registry): stamp the
# fixtures with a neutral name, then put the user's own name back in the finally block.
$previousUserName = $excel.UserName
$excel.UserName = 'xlsx-viewer'
try {
	$book = $excel.Workbooks.Open($sample)
	$excel.CalculateFull()
	$book.BuiltinDocumentProperties.Item('Last Author').Value = 'xlsx-viewer'
	$book.Worksheets.Item(1).Activate()
	$book.SaveAs($sample, $xlOpenXMLWorkbook)
	$book.SaveAs($legacy, $xlExcel8)
	Write-Output "sheets: $($book.Worksheets.Count)"
	$book.Close($false)
	$check = $excel.Workbooks.Open($legacy)
	Write-Output "reopened $legacy with $($check.Worksheets.Count) sheets"
	$check.Close($false)
} finally {
	$excel.UserName = $previousUserName
	$excel.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "wrote $sample and $legacy"
