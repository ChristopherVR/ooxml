# Requires desktop Excel and Python with PyMuPDF. Files remain temporary.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskFolder = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-print-format-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
$previousRoot=$env:OOXML_PRINT_ROOT; $previousOutput=$env:OOXML_PRINT_OUTPUT
$excel=$null; $book=$null
try {
	$env:OOXML_PRINT_ROOT=$taskRoot; $env:OOXML_PRINT_OUTPUT=$taskFolder
	$script=Join-Path $taskFolder 'generate.ts'
	@'
import {writeFileSync,readFileSync} from 'node:fs';import {join} from 'node:path';import {pathToFileURL} from 'node:url';
const root=process.env.OOXML_PRINT_ROOT!;const out=process.env.OOXML_PRINT_OUTPUT!;
const {createWorkbook,createEditSession,saveXlsx,loadXlsx}=await import(pathToFileURL(join(root,'src/core/xlsx/index.ts')).href);
const header='&C&"Arial,Bold Italic"&18&KFF0000Rich &P&B Normal';
const wb=createWorkbook();const edit=createEditSession(wb);edit.setCellInput(0,0,0,'Cell');
edit.setPageSetup(0,{header,footer:'&C&"Arial,Regular"&12&UUnder&U &SStrike&S',scale:100});
const bytes=await saveXlsx(wb);const loaded=await loadXlsx(bytes);
if(loaded.sheets[0].pageSetup.header!==header)throw Error('Rich header round trip changed');
writeFileSync(join(out,'rich.xlsx'),bytes);
'@ | Set-Content -LiteralPath $script
	& bun run $script
	if($LASTEXITCODE -ne 0){throw 'Rich print workbook generation failed'}
	$excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
	Write-Output ('Native Excel version: '+(Get-Item -LiteralPath (Join-Path $excel.Path 'EXCEL.EXE')).VersionInfo.FileVersion)
	$book=$excel.Workbooks.Open((Join-Path $taskFolder 'rich.xlsx'))
	$sheet=$book.Worksheets.Item(1)
	if($sheet.PageSetup.CenterHeader -ne '&"Arial,Bold Italic"&18&KFF0000Rich &P&B Normal'){throw 'Native rich header differs'}
	$sheet.ExportAsFixedFormat(0,(Join-Path $taskFolder 'rich.pdf'))
	$book.SaveAs((Join-Path $taskFolder 'native-saved.xlsx'),51)
	$book.Close($false); $book=$null
	$probe=Join-Path $taskFolder 'probe.py'
	@'
import fitz,sys
doc=fitz.open(sys.argv[1])
spans=[s for page in doc for b in page.get_text('dict')['blocks'] if 'lines' in b for line in b['lines'] for s in line['spans']]
rich=next(s for s in spans if 'Rich' in s['text'])
normal=next(s for s in spans if 'Normal' in s['text'])
assert abs(rich['size']-18)<0.1 and rich['color']==0xff0000,rich
assert 'Bold' in rich['font'] and ('Italic' in rich['font'] or 'Oblique' in rich['font']),rich
assert 'Bold' not in normal['font'] and normal['color']==0xff0000,normal
assert 'Under' in ''.join(s['text'] for s in spans) and 'Strike' in ''.join(s['text'] for s in spans)
print('Native PDF: 18pt red bold italic header; bold toggled off; underline/strike text accepted')
print(rich)
'@ | Set-Content -LiteralPath $probe
	& python -I $probe (Join-Path $taskFolder 'rich.pdf')
	if($LASTEXITCODE -ne 0){throw 'Native rich PDF formatting differs'}
	@'
import {readFileSync,writeFileSync} from 'node:fs';import {join} from 'node:path';import {pathToFileURL} from 'node:url';
const root=process.env.OOXML_PRINT_ROOT!;const out=process.env.OOXML_PRINT_OUTPUT!;
const {loadXlsx,saveXlsx,renderHeaderFooterRuns}=await import(pathToFileURL(join(root,'src/core/xlsx/index.ts')).href);
const wb=await loadXlsx(readFileSync(join(out,'native-saved.xlsx')));
const result=renderHeaderFooterRuns(wb.sheets[0].pageSetup.header,{page:1,pages:1,sheet:'Sheet1',file:'rich.xlsx'});
if(result.unsupported.length||!result.sections.center.some(r=>r.style.color==='#FF0000'&&r.style.bold&&r.style.size===18))throw Error('Native rich formatting not parsed');
writeFileSync(join(out,'roundtrip.xlsx'),await saveXlsx(wb));
'@ | Set-Content -LiteralPath $script
	& bun run $script
	if($LASTEXITCODE -ne 0){throw 'Native rich round trip failed'}
	$book=$excel.Workbooks.Open((Join-Path $taskFolder 'roundtrip.xlsx'))
	if($book.Worksheets.Item(1).PageSetup.CenterHeader -notmatch 'Rich &P'){throw 'Native reopen lost header'}
	Write-Output 'Native save/import/export/reopen passed'
}finally{
	if($book){$book.Close($false)}
	if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
	$env:OOXML_PRINT_ROOT=$previousRoot; $env:OOXML_PRINT_OUTPUT=$previousOutput
}
Write-Output ('Temporary acceptance output: '+$taskFolder)
