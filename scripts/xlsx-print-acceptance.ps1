# Requires desktop Excel, pdfinfo, and Python with pypdf. Output remains temporary.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskFolder = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-print-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
$previousRoot = $env:OOXML_PRINT_ROOT; $previousOutput = $env:OOXML_PRINT_OUTPUT
$excel = $null; $book = $null
try {
	$env:OOXML_PRINT_ROOT = $taskRoot; $env:OOXML_PRINT_OUTPUT = $taskFolder
	$script = Join-Path $taskFolder 'generate.ts'
	@'
import { writeFileSync } from 'node:fs'; import { join } from 'node:path'; import { pathToFileURL } from 'node:url';
const root=process.env.OOXML_PRINT_ROOT!; const out=process.env.OOXML_PRINT_OUTPUT!;
const {createWorkbook,createEditSession,saveXlsx,parseRange}=await import(pathToFileURL(join(root,'src/core/xlsx/index.ts')).href);
const book=createWorkbook(); const edit=createEditSession(book);
for(let r=0;r<10;r++) for(let c=0;c<4;c++) edit.setCellInput(0,r,c,`R${r+1}C${c+1}`);
edit.setPageSetup(0,{orientation:'landscape',paperSize:9,scale:75,pageOrder:'overThenDown',printArea:parseRange('A1:D10'),margins:{left:.25,right:.25,top:.25,bottom:.25,header:.1,footer:.1}});
edit.setPrintOptions(0,{gridLines:true,headings:true}); edit.setPageBreak(0,'row',5); edit.setPageBreak(0,'col',2);
edit.setPrintTitles(0,{rows:{start:0,end:0},columns:{start:0,end:0}});
edit.setPageSetup(0,{firstPageNumber:5,header:'&L&A&CReport&RPage &P of &N',footer:'&L&F&C&D &T',differentFirst:true,firstHeader:'&CFirst &P of &N',firstFooter:'&CFIRST',differentOddEven:true,evenHeader:'&CEven &P of &N',evenFooter:'&L&F&R&D &T'});
writeFileSync(join(out,'authored.xlsx'),await saveXlsx(book));
edit.resetPageBreaks(0);writeFileSync(join(out,'reset.xlsx'),await saveXlsx(book));
edit.undo(); edit.setPageSetup(0,{scale:undefined,fitToWidth:1,fitToHeight:0});writeFileSync(join(out,'fit.xlsx'),await saveXlsx(book));
'@ | Set-Content -LiteralPath $script
	& bun run $script
	if ($LASTEXITCODE -ne 0) { throw 'Print workbook generation failed.' }
	$excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false
	Write-Output ('Native Excel version: '+(Get-Item -LiteralPath (Join-Path $excel.Path 'EXCEL.EXE')).VersionInfo.FileVersion)
	foreach ($case in @('authored','reset','fit')) {
		$book = $excel.Workbooks.Open((Join-Path $taskFolder ($case+'.xlsx')))
		$sheet = $book.Worksheets.Item(1); $setup = $sheet.PageSetup
		if ($setup.Orientation -ne 2 -or $setup.PaperSize -ne 9 -or $setup.Order -ne 2 -or [Math]::Abs($setup.LeftMargin - 18) -gt .01) { throw 'Native paper, orientation, order or margins differ.' }
		if (-not $setup.PrintGridlines -or -not $setup.PrintHeadings) { throw 'Native print options differ.' }
		if ($setup.PrintTitleRows -ne '$1:$1' -or $setup.PrintTitleColumns -ne '$A:$A' -or $setup.FirstPageNumber -ne 5) { throw 'Native title ranges or first page number differ.' }
		if ($setup.LeftHeader -ne '&A' -or $setup.RightHeader -ne 'Page &P of &N' -or -not $setup.DifferentFirstPageHeaderFooter -or -not $setup.OddAndEvenPagesHeaderFooter) { throw 'Native header fields or variants differ.' }
		if ($case -eq 'authored') {
			if ($setup.Zoom -ne 75) { throw 'Native scale differs.' }
			if ($sheet.HPageBreaks().Count -ne 1 -or $sheet.HPageBreaks().Item(1).Location.Row -ne 6) { throw 'Native row page break differs.' }
			if ($sheet.VPageBreaks().Count -ne 1 -or $sheet.VPageBreaks().Item(1).Location.Column -ne 3) { throw 'Native column page break differs.' }
			$pdf = Join-Path $taskFolder 'native-print.pdf'; $sheet.ExportAsFixedFormat(0,$pdf)
			$info = & pdfinfo $pdf
			if (-not ($info -match '^Pages:\s+4\s*$')) { throw 'Native printed page count is not four.' }
			$printed = (& python -c "import sys; from pypdf import PdfReader; print('\n'.join(p.extract_text() or '' for p in PdfReader(sys.argv[1]).pages))" $pdf) -join "`n"
			if ($printed -notmatch 'First\s+5\s+of\s+4' -or $printed -notmatch 'Even\s+6\s+of\s+4' -or $printed -notmatch 'Page\s+7\s+of\s+4' -or $printed -notmatch 'authored.xlsx') { throw 'Native header field text or page variants differ.' }
			if ($printed -notmatch '(?:\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4})\s+\d{1,2}:\d{2}') { throw 'Native date/time fields differ.' }
			if ([regex]::Matches($printed, 'R1C1').Count -ne 4) { throw 'Native row and column titles did not repeat on each page.' }
			$sheet.HPageBreaks().Add($sheet.Range('A8')) | Out-Null
			$book.SaveAs((Join-Path $taskFolder 'native-edited.xlsx'),51)
		} elseif ($case -eq 'reset') {
			if ($sheet.HPageBreaks().Count -ne 0 -or $sheet.VPageBreaks().Count -ne 0) { throw 'Native reset kept manual breaks.' }
		} else {
			if ($setup.Zoom -ne $false -or $setup.FitToPagesWide -ne 1 -or $setup.FitToPagesTall -ne $false) { throw 'Native fit-to-page settings differ.' }
		}
		Write-Output ($case+': native Excel accepted print settings and page breaks')
		$book.Close($false); $book = $null
	}
	@'
import {readFileSync,writeFileSync} from 'node:fs';import {join} from 'node:path';import {pathToFileURL} from 'node:url';
const root=process.env.OOXML_PRINT_ROOT!;const out=process.env.OOXML_PRINT_OUTPUT!;
const {loadXlsx,saveXlsx,createEditSession,sheetPageBreaks}=await import(pathToFileURL(join(root,'src/core/xlsx/index.ts')).href);
const book=await loadXlsx(new Uint8Array(readFileSync(join(out,'native-edited.xlsx'))));
if(!sheetPageBreaks(book.sheets[0],'row').some(b=>b.before===7&&b.manual))throw Error('Native-authored break was not parsed');
createEditSession(book).setPageBreak(0,'row',7,false);writeFileSync(join(out,'native-roundtrip.xlsx'),await saveXlsx(book));
'@ | Set-Content -LiteralPath $script
	& bun run $script
	if ($LASTEXITCODE -ne 0) { throw 'Native print round trip failed.' }
	$book = $excel.Workbooks.Open((Join-Path $taskFolder 'native-roundtrip.xlsx'))
	if ($book.Worksheets.Item(1).HPageBreaks().Count -ne 1) { throw 'Removed native break survived round trip.' }
	Write-Output 'native-roundtrip: Excel-authored break removed, saved and reopened'
	$book.Close($false); $book = $null
} finally {
	if ($book) { $book.Close($false) }
	if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
	$env:OOXML_PRINT_ROOT=$previousRoot; $env:OOXML_PRINT_OUTPUT=$previousOutput
}
Write-Output ('Temporary acceptance output: '+$taskFolder)
