# Manual acceptance against installed desktop Excel for slicers, timelines and number groups.
# Writes workbooks with the core (inserted, filtered and edited Excel-authored files), opens each in
# Excel and compares slicer item states and PivotTable text with what the core computed.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskFolder = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-slicers-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
$generator = Join-Path $taskFolder 'generate.ts'
$previousRoot = $env:OOXML_SLICER_ROOT
$previousOutput = $env:OOXML_SLICER_OUTPUT
$excel = $null; $book = $null
try {
	$env:OOXML_SLICER_ROOT = $taskRoot; $env:OOXML_SLICER_OUTPUT = $taskFolder
	@'
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
const root = process.env.OOXML_SLICER_ROOT!; const out = process.env.OOXML_SLICER_OUTPUT!;
const { createWorkbook, createEditSession, saveXlsx, loadXlsx, slicerItems } = await import(pathToFileURL(join(root, 'src/core/xlsx/index.ts')).href);
const fixtures = join(root, 'src/core/xlsx/__fixtures__');
const read = (name: string) => new Uint8Array(readFileSync(join(fixtures, name)));
const cases: Record<string, unknown> = {};
const finish = async (name: string, wb: any, pivots: Record<string, string[][]> = {}, slicers: Record<string, string[]> = {}) => {
	writeFileSync(join(out, name + '.xlsx'), await saveXlsx(wb));
	cases[name] = { pivots, slicers };
};
const keyOf = (wb: any, cache: any, label: string) => slicerItems(wb, cache).find((item: any) => item.label === label).key;

// Table slicer: filter an Excel-authored table slicer.
{
	const wb = await loadXlsx(read('excel-slicer-table.xlsx'));
	createEditSession(wb).setSlicerSelection('Slicer_Category', ['Veg']);
	await finish('table-filter', wb, {}, { Slicer_Category: ['Fruit=False', 'Nuts=False', 'Veg=True'], Slicer_Region: ['East=True', 'West=True'] });
}
// PivotTable slicer: hide Fruit on an Excel-authored slicer-connected PivotTable.
{
	const wb = await loadXlsx(read('excel-slicer-pivot.xlsx'));
	const cache = wb.slicers.caches[0];
	createEditSession(wb).setSlicerSelection(cache.name, [keyOf(wb, cache, 'Nuts'), keyOf(wb, cache, 'Veg')]);
	await finish('pivot-filter', wb, { SalesPivot: [['Sum of Amount', 'Column Labels', '', ''], ['Row Labels', 'East', 'West', 'Grand Total'], ['Nuts', '', '9', '9'], ['Veg', '5', '15', '20'], ['Grand Total', '5', '24', '29']] }, { Slicer_Category: ['Fruit=False', 'Nuts=True', 'Veg=True'] });
}
// Inserted table and PivotTable slicers on a core-authored workbook.
{
	const wb = createWorkbook({ sheets: ['Data', 'Summary'] }); const edit = createEditSession(wb);
	const rows = [['Category', 'Region', 'Amount'], ['Fruit', 'East', '10'], ['Fruit', 'West', '20'], ['Veg', 'East', '5'], ['Veg', 'West', '15']];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createTable(0, { start: { row: 0, col: 0 }, end: { row: 4, col: 2 } }, true);
	edit.createPivotTable(1, { row: 0, col: 0 }, { source: 'Data!A1:C5', rowField: 'Category', valueField: 'Amount' });
	edit.insertSlicer(0, { kind: 'table', table: 'Table1', column: 'Region' });
	edit.insertSlicer(1, { kind: 'pivot', pivot: 'PivotTable1', field: 'Category' });
	edit.setSlicerSelection('Slicer_Region', ['west']);
	edit.setSlicerSelection('Slicer_Category', ['string:veg']);
	await finish('inserted', wb, { PivotTable1: [['Category', 'Sum of Amount'], ['Veg', '20'], ['Grand Total', '20']] }, { Slicer_Region: ['East=False', 'West=True'], Slicer_Category: ['Fruit=False', 'Veg=True'] });
}
// Timeline: select Feb to Apr on an Excel-authored timeline, then a new timeline.
{
	const wb = await loadXlsx(read('excel-slicer-timeline.xlsx'));
	createEditSession(wb).setTimelineSelection('Day', { start: '2025-02-01T00:00:00', end: '2025-04-30T00:00:00' });
	await finish('timeline-edit', wb, { SalesPivot: [['Row Labels', 'Sum of Amount'], ['Fruit', '20'], ['Veg', '20'], ['Grand Total', '40']] });
	const nw = createWorkbook({ sheets: ['Data', 'Summary'] }); const edit = createEditSession(nw);
	const rows = [['Category', 'Amount', 'Day'], ['Fruit', '10', '2025-01-15'], ['Fruit', '20', '2025-02-15'], ['Veg', '5', '2025-03-15']];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createPivotTable(1, { row: 0, col: 0 }, { source: 'Data!A1:C4', rowField: 'Category', valueField: 'Amount' });
	edit.insertTimeline(1, { pivot: 'PivotTable1', field: 'Day' });
	edit.setTimelineSelection('NativeTimeline_Day', { start: '2025-02-01T00:00:00', end: '2025-03-31T00:00:00' });
	await finish('timeline-new', nw, { PivotTable1: [['Category', 'Sum of Amount'], ['Fruit', '20'], ['Veg', '5'], ['Grand Total', '25']] });
}
// Number groups: regroup an Excel-authored group, and group a new report.
{
	const wb = await loadXlsx(read('excel-pivot-group-number.xlsx'));
	createEditSession(wb).refreshPivotTable(0, 'GroupPivot', { fieldGroups: [{ field: 'Score', start: 0, end: 60, by: 30 }] });
	await finish('group-edit', wb, { GroupPivot: [['Row Labels', 'Sum of Amount'], ['0-29', '57'], ['30-60', '9'], ['Grand Total', '66']] });
	const nw = createWorkbook({ sheets: ['Data', 'Summary'] }); const edit = createEditSession(nw);
	const rows = [['Category', 'Score', 'Amount'], ['A', '3', '10'], ['B', '12', '20'], ['A', '61', '4']];
	rows.forEach((line, r) => line.forEach((value, c) => edit.setCellInput(0, r, c, value)));
	edit.createPivotTable(1, { row: 0, col: 0 }, { source: 'Data!A1:C4', rowField: 'Score', valueField: 'Amount' });
	edit.refreshPivotTable(1, 'PivotTable1', { fieldGroups: [{ field: 'Score', start: 10, end: 50, by: 10 }] });
	await finish('group-new', nw, { PivotTable1: [['Score', 'Sum of Amount'], ['<10', '10'], ['10-19', '20'], ['>50', '4'], ['Grand Total', '34']] });
}
writeFileSync(join(out, 'cases.json'), JSON.stringify(cases));
'@ | Set-Content -Path $generator -Encoding UTF8
	Push-Location $taskRoot
	try { bun run $generator } finally { Pop-Location }
	$cases = Get-Content -Raw (Join-Path $taskFolder 'cases.json') | ConvertFrom-Json
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false; $excel.DisplayAlerts = $false
	$failures = 0
	foreach ($case in $cases.PSObject.Properties) {
		$path = Join-Path $taskFolder ($case.Name + '.xlsx')
		try { $book = $excel.Workbooks.Open($path) } catch { Write-Host "FAIL $($case.Name): Excel could not open the file"; $failures++; continue }
		$problems = @()
		foreach ($slicer in $case.Value.slicers.PSObject.Properties) {
			$cache = $null; foreach ($c in $book.SlicerCaches) { if ($c.Name -eq $slicer.Name) { $cache = $c } }
			if (-not $cache) { $problems += "missing slicer cache $($slicer.Name)"; continue }
			$states = @(); foreach ($item in $cache.SlicerItems) { $states += ($item.Name + '=' + $item.Selected) }
			if (($states -join ',') -ne (@($slicer.Value) -join ',')) { $problems += "slicer $($slicer.Name): $($states -join ',')" }
		}
		foreach ($pivotCase in $case.Value.pivots.PSObject.Properties) {
			$found = $null
			foreach ($sheet in $book.Worksheets) { $n = $sheet.PivotTables().Count; for ($k = 1; $k -le $n; $k++) { if ($sheet.PivotTables($k).Name -eq $pivotCase.Name) { $found = $sheet.PivotTables($k) } } }
			if (-not $found) { $problems += "missing PivotTable $($pivotCase.Name)"; continue }
			$range = $found.TableRange2; $expected = @($pivotCase.Value)
			if ($range.Rows.Count -ne $expected.Count) { $problems += "pivot $($pivotCase.Name): $($range.Rows.Count) rows, expected $($expected.Count)"; continue }
			for ($r = 1; $r -le $range.Rows.Count; $r++) {
				for ($c = 1; $c -le $expected[$r - 1].Count; $c++) {
					$text = [string]$range.Cells($r, $c).Text
					if ($text -ne [string]$expected[$r - 1][$c - 1]) { $problems += "pivot $($pivotCase.Name) r$r c${c}: '$text' vs '$($expected[$r - 1][$c - 1])'" }
				}
			}
		}
		$book.Close($false); $book = $null
		if ($problems.Count) { Write-Host "FAIL $($case.Name): $($problems -join '; ')"; $failures++ } else { Write-Host "ok   $($case.Name)" }
	}
	if ($failures) { throw "$failures native slicer acceptance case(s) failed." }
	Write-Host 'All native slicer, timeline and group cases passed.'
} finally {
	if ($book) { $book.Close($false) }
	if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
	$env:OOXML_SLICER_ROOT = $previousRoot; $env:OOXML_SLICER_OUTPUT = $previousOutput
}
