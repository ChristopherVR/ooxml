# Manual acceptance against installed desktop Excel. Generates temporary workbooks with our
# writer, checks native sparkline types/settings, and extends an Excel-authored fixture.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskFolder = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-sparklines-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
$taskScript = Join-Path $taskFolder 'generate.ts'
$taskPreviousRoot = $env:OOXML_SPARKLINE_ROOT
$taskPreviousOutput = $env:OOXML_SPARKLINE_OUTPUT
$taskExcel = $null
$taskBook = $null
try {
	$env:OOXML_SPARKLINE_ROOT = $taskRoot
	$env:OOXML_SPARKLINE_OUTPUT = $taskFolder
	@'
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
const root = process.env.OOXML_SPARKLINE_ROOT!;
const out = process.env.OOXML_SPARKLINE_OUTPUT!;
const { createWorkbook, createEditSession, parseRange, saveXlsx, loadXlsx } = await import(pathToFileURL(join(root, 'src/core/xlsx/index.ts')).href);
const wb = createWorkbook({ sheets: ['Sales'] });
const edit = createEditSession(wb);
for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) edit.setCellInput(0, r, c, String((r + 1) * (c === 1 ? -3 : c + 2)));
edit.createSparklineGroup(0, 'A1:C2', parseRange('D1:D2')!, { type: 'column', high: true, negative: true, minAxisType: 'custom', manualMin: -10, maxAxisType: 'custom', manualMax: 10, colorSeries: { rgb: '123456' } });
edit.createSparklineGroup(0, 'A1:C1', parseRange('E1')!, { type: 'line', markers: true });
edit.createSparklineGroup(0, 'A2:C2', parseRange('E2')!, { type: 'stacked', negative: true });
for (let c = 0; c < 3; c++) edit.setCellInput(0, 3, c, String([46000, 46001, 46020][c]));
edit.setSparklineDateRange(0, [parseRange('E1')!], 'A4:C4');
writeFileSync(join(out, 'authored.xlsx'), await saveXlsx(wb));
edit.setSparklineDateRange(0, [parseRange('E1')!]);
writeFileSync(join(out, 'date-cleared.xlsx'), await saveXlsx(wb));
edit.undo();
writeFileSync(join(out, 'date-undone.xlsx'), await saveXlsx(wb));
edit.deleteSparklines(0, [parseRange('D1:D2')!, parseRange('E1:E2')!], true);
writeFileSync(join(out, 'cleared.xlsx'), await saveXlsx(wb));
edit.undo();
writeFileSync(join(out, 'undone.xlsx'), await saveXlsx(wb));
const imported = await loadXlsx(new Uint8Array(readFileSync(join(root, 'src/core/xlsx/__fixtures__/excel-sparklines.xlsx'))));
createEditSession(imported).createSparklineGroup(0, 'A1:E1', parseRange('G1')!, { type: 'line' });
writeFileSync(join(out, 'imported.xlsx'), await saveXlsx(imported));
'@ | Set-Content -LiteralPath $taskScript
	& bun run $taskScript
	if ($LASTEXITCODE -ne 0) { throw 'Sparkline workbook generation failed.' }
	$taskExcel = New-Object -ComObject Excel.Application
	$taskExcel.Visible = $false
	$taskExcel.DisplayAlerts = $false
	Write-Output ('Native Excel version: ' + $taskExcel.Version + ', build: ' + $taskExcel.Build)
	foreach ($taskCase in @(@{ Name = 'authored'; Count = 3 }, @{ Name = 'cleared'; Count = 0 }, @{ Name = 'undone'; Count = 3 }, @{ Name = 'imported'; Count = 5 })) {
		$taskBook = $taskExcel.Workbooks.Open((Join-Path $taskFolder ($taskCase.Name + '.xlsx')))
		$taskSheet = $taskBook.Worksheets.Item(1)
		if ($taskSheet.Cells.SparklineGroups.Count -ne $taskCase.Count) { throw ('Wrong native group count for ' + $taskCase.Name) }
		if ($taskCase.Name -eq 'authored') {
			$taskGroup = $taskSheet.Range('D1').SparklineGroups.Item(1)
			if ($taskGroup.Type -ne 2 -or $taskGroup.SourceData -ne 'A1:C2' -or -not $taskGroup.Points.Highpoint.Visible) { throw 'Native column sparkline settings did not survive.' }
			if ($taskGroup.Axes.Vertical.CustomMinScaleValue -ne -10 -or $taskGroup.Axes.Vertical.CustomMaxScaleValue -ne 10) { throw 'Native custom sparkline scale did not survive.' }
			if ($taskGroup.SeriesColor.Color -ne 0x563412) { throw 'Native sparkline RGB color did not survive.' }
			if ($taskSheet.Range('E1').SparklineGroups.Item(1).Type -ne 1 -or $taskSheet.Range('E2').SparklineGroups.Item(1).Type -ne 3) { throw 'Native line or win/loss type did not survive.' }
		}
		Write-Output ($taskCase.Name + ': native Excel accepted ' + $taskCase.Count + ' sparkline groups')
		$taskBook.Close($false)
		$taskBook = $null
	}
	foreach ($taskCase in @(@{ Name = 'authored'; Date = 'A4:C4' }, @{ Name = 'date-cleared'; Date = '' }, @{ Name = 'date-undone'; Date = 'A4:C4' })) {
		$taskBook = $taskExcel.Workbooks.Open((Join-Path $taskFolder ($taskCase.Name + '.xlsx')))
		$taskDate = $taskBook.Worksheets.Item(1).Range('E1').SparklineGroups.Item(1).DateRange
		if ([string]$taskDate -ne $taskCase.Date) { throw ('Wrong native date range for ' + $taskCase.Name + ': ' + $taskDate) }
		Write-Output ($taskCase.Name + ': native date range accepted: ' + $taskDate)
		$taskBook.Close($false)
		$taskBook = $null
	}
} finally {
	if ($taskBook) { $taskBook.Close($false) }
	if ($taskExcel) { $taskExcel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskExcel) }
	$env:OOXML_SPARKLINE_ROOT = $taskPreviousRoot
	$env:OOXML_SPARKLINE_OUTPUT = $taskPreviousOutput
	# Keep the generated files for inspection; do not remove an arbitrary computed directory.
	Write-Output ('Acceptance files: ' + $taskFolder)
}
