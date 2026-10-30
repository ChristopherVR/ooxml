# Generates a native x14 cross-sheet list validation through Excel 16.
$ErrorActionPreference = 'Stop'
$validationExcel = $null
$validationBook = $null
try {
	$validationExcel = New-Object -ComObject Excel.Application
	$validationExcel.Visible = $false
	$validationExcel.DisplayAlerts = $false
	$validationBook = $validationExcel.Workbooks.Add()
	$inputSheet = $validationBook.Worksheets.Item(1)
	$inputSheet.Name = 'Input'
	$lookupSheet = $validationBook.Worksheets.Add()
	$lookupSheet.Name = 'Lookup'
	$lookupSheet.Range('A1').Value2 = 'red'
	$lookupSheet.Range('A2').Value2 = 'green'
	$lookupSheet.Range('A3').Value2 = 'blue'
	$inputSheet.Range('B2:B5').Validation.Add(3, 1, 1, '=Lookup!$A$1:$A$3')
	$inputSheet.Range('B2:B5').Validation.ErrorTitle = 'Pick a colour'
	$inputSheet.Range('B2:B5').Validation.ErrorMessage = 'Choose from the list.'
	$inputSheet.Range('B2:B5').Validation.ShowError = $true
	$inputSheet.Activate()
	$validationBook.SaveAs((Join-Path $PSScriptRoot 'excel-x14-validation.xlsx'), 51)
} finally {
	if ($null -ne $validationBook) {
		$validationBook.Close($false)
		[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($validationBook)
	}
	if ($null -ne $validationExcel) {
		$validationExcel.Quit()
		[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($validationExcel)
	}
}
