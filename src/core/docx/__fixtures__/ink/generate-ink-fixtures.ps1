# Generates the ink fixtures of Word and Excel from PowerPoint's own ink (the committed
# `pptx/__tests__/fixtures/e2e/ink-contentpart.pptx`), so every file is the applications' own save:
#   word-ink-floating.docx    the ink pasted into Word (floating, wrapped in mc:AlternateContent with
#                             a picture fallback), saved by Word.
#   ../../xlsx/__fixtures__/excel-ink-fallback.xlsx   the ink pasted into Excel (with its fallback).
# word-ink-inline.docx and excel-ink.xlsx are Word's and Excel's saves of the same InkML part placed
# by hand-written markup (an inline `w14:contentPart`, and an `xdr:contentPart` with no fallback):
# the applications rewrote both in their own form when they saved them.
# Requires desktop PowerPoint, Word and Excel (COM automation). Run from any folder:
#   pwsh -File src/core/docx/__fixtures__/ink/generate-ink-fixtures.ps1
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$deck = Resolve-Path (Join-Path $here '../../../pptx/__tests__/fixtures/e2e/ink-contentpart.pptx')
$powerpoint = New-Object -ComObject PowerPoint.Application
$presentation = $powerpoint.Presentations.Open($deck, $true, $false, $false)
try {
	$ink = $presentation.Slides.Item(1).Shapes | Where-Object { $_.Type -eq 23 } | Select-Object -First 1
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false
	$excel.DisplayAlerts = $false
	try {
		$ink.Copy()
		$book = $excel.Workbooks.Add()
		$sheet = $book.Worksheets.Item(1)
		$sheet.Range('A1').Value2 = 'Ink'
		$sheet.Range('B2').Select() | Out-Null
		$sheet.Paste()
		$book.SaveAs((Join-Path $here '../../../xlsx/__fixtures__/excel-ink-fallback.xlsx'), 51)
		$book.Close($false)
	} finally { $excel.Quit() }
	$word = New-Object -ComObject Word.Application
	$word.Visible = $false
	$word.DisplayAlerts = 0
	try {
		$ink.Copy()
		$document = $word.Documents.Add()
		$document.Range().Paste()
		$document.SaveAs2((Join-Path $here 'word-ink-floating.docx'), 16)
		$document.Close($false)
	} finally { $word.Quit() }
} finally {
	$presentation.Close()
	$powerpoint.Quit()
}
