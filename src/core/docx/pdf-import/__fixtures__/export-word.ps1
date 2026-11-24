# Exports each source document to PDF with an owned Microsoft Word 16 through its documented
# COM object model (Documents.Open, ExportAsFixedFormat). Run on Windows with Word installed.
# Pass base names (`-Names tableaux,listes`) to export only those documents.
param([string[]]$Names)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
	foreach ($file in Get-ChildItem (Join-Path $here 'source') -Filter *.docx) {
		if ($Names -and $Names -notcontains $file.BaseName) { continue }
		$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
		$pdf = Join-Path $here "$($file.BaseName).pdf"
		# 17 = wdExportFormatPDF; document properties and structure tags are included.
		$doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 1, 1, 0, $true, $true, 0, $true, $true, $false)
		"$($file.BaseName): $($doc.ComputeStatistics(2)) pages"
		$doc.Close(0)
	}
} finally { $word.Quit() }
