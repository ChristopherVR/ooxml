# Manual acceptance against installed desktop Excel; outputs stay in a temporary folder.
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskFolder = Join-Path ([IO.Path]::GetTempPath()) ('ooxml-pivots-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $taskFolder | Out-Null
$taskScript = Join-Path $taskFolder 'generate.ts'
$previousRoot = $env:OOXML_PIVOT_ROOT
$previousOutput = $env:OOXML_PIVOT_OUTPUT
$excel = $null; $book = $null
try {
	$env:OOXML_PIVOT_ROOT = $taskRoot; $env:OOXML_PIVOT_OUTPUT = $taskFolder
	@'
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
const root = process.env.OOXML_PIVOT_ROOT!; const out = process.env.OOXML_PIVOT_OUTPUT!;
const { createWorkbook, createEditSession, saveXlsx, loadXlsx, pivotGrid } = await import(pathToFileURL(join(root, 'src/core/xlsx/index.ts')).href);
for (const aggregate of ['sum', 'count', 'average', 'min', 'max']) {
 const wb = createWorkbook({sheets:['Data','Summary']}); const edit = createEditSession(wb);
 const rows = [['Category','Region','Amount'],['Fruit','East',10],['Fruit','West',20],['Veg','East',5],['Veg','West',15],['Fruit','East',7]];
 rows.forEach((line,r)=>line.forEach((value,c)=>edit.setCellInput(0,r,c,String(value))));
 edit.createPivotTable(1,{row:0,col:0},{source:'Data!A1:C6',rowField:'Category',columnField:'Region',valueField:'Amount',aggregate});
 writeFileSync(join(out,aggregate+'.xlsx'),await saveXlsx(wb));
}
const imported = await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot.xlsx'))));
const edit = createEditSession(imported); edit.setCellInput(1,1,2,'100'); edit.refreshPivotTable(0,'SalesPivot');
writeFileSync(join(out,'imported.xlsx'),await saveXlsx(imported));
edit.refreshPivotTable(0,'SalesPivot',{columnField:'',valueField:'Category',aggregate:'count'});
writeFileSync(join(out,'count-row.xlsx'),await saveXlsx(imported));
edit.deletePivotTable(0,'SalesPivot'); writeFileSync(join(out,'deleted.xlsx'),await saveXlsx(imported));
for (const mode of ['cache-blank', 'cache-numeric', 'cache-boolean', 'cache-error']) {
 const wb=createWorkbook({sheets:['Data','Summary']});const session=createEditSession(wb);
 const rows:(string|number|boolean)[][]=[['Category','Region','Amount'],['Fruit','East',10],['Fruit','West',20],['Veg','East',5],['Veg','West',15],['Fruit','East',7]];
 for(let r=1;r<rows.length;r++) {
  if(mode==='cache-blank' && rows[r][0]==='Fruit') rows[r][0]='';
  if(mode==='cache-numeric') rows[r][0]=rows[r][0]==='Fruit'?2026:2027;
  if(mode==='cache-boolean') rows[r][0]=rows[r][0]==='Fruit';
  if(mode==='cache-error' && rows[r][0]==='Fruit') rows[r][0]='=NA()';
 }
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:C6',rowField:'Category',columnField:'Region',valueField:'Amount'});
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[1].pivotTables[0]).values));
}
'@ | Set-Content -LiteralPath $taskScript
	& bun run $taskScript
	if ($LASTEXITCODE -ne 0) { throw 'PivotTable generation failed.' }
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false; $excel.DisplayAlerts = $false
	Write-Output ('Native Excel version: '+$excel.Version+', build: '+$excel.Build)
	foreach ($case in @(@{ Name='sum'; Total=57 }, @{ Name='count'; Total=5 }, @{ Name='average'; Total=11.4 }, @{ Name='min'; Total=5 }, @{ Name='max'; Total=20 }, @{ Name='imported'; Total=147 }, @{ Name='count-row'; Total=5 }, @{ Name='deleted'; Total=$null })) {
		$book = $excel.Workbooks.Open((Join-Path $taskFolder ($case.Name+'.xlsx')))
		$summary = $book.Worksheets.Item('Summary')
		if ($case.Name -eq 'deleted') {
			if ($summary.PivotTables().Count -ne 0) { throw 'Deleted PivotTable still present.' }
		} else {
			if ($summary.PivotTables().Count -ne 1) { throw ('Missing native PivotTable: '+$case.Name) }
			$cell = if ($case.Name -eq 'count-row') { 'B4' } else { 'D5' }
			if ([Math]::Abs([double]$summary.Range($cell).Value2 - $case.Total) -gt 0.00001) { throw ('Incorrect cached total: '+$case.Name) }
			$pivot = $summary.PivotTables().Item(1)
			[void]$pivot.RefreshTable()
			if ([Math]::Abs([double]$summary.Range($cell).Value2 - $case.Total) -gt 0.00001) { throw ('Incorrect refreshed total: '+$case.Name) }
			if ($case.Name -eq 'sum') {
				$book.Worksheets.Item('Data').Range('C2').Value2 = [double]100
				[void]$pivot.RefreshTable()
				if ($summary.Range('D5').Value2 -ne 147) { throw 'Native refresh did not read changed source.' }
			}
		}
		Write-Output ($case.Name+': native Excel accepted cached and refreshed output')
		$book.Close($false); $book = $null
	}
	foreach ($mode in @('cache-blank','cache-numeric','cache-boolean','cache-error')) {
		$book = $excel.Workbooks.Open((Join-Path $taskFolder ($mode+'.xlsx')))
		$summary = $book.Worksheets.Item('Summary')
		$pivot = $summary.PivotTables().Item(1)
		if ($pivot.RowFields().Count -ne 1) { throw ('Wrong PivotTable row field count: '+$mode) }
		$expected = Get-Content -Raw -LiteralPath (Join-Path $taskFolder ($mode+'.json')) | ConvertFrom-Json
		foreach ($phase in @('cached','refreshed')) {
			if ($phase -eq 'refreshed') { [void]$pivot.RefreshTable() }
			if ($pivot.TableRange2.Rows.Count -ne $expected.Count) { throw ('Wrong cache row count: '+$mode+' '+$phase) }
			if ($pivot.TableRange2.Columns.Count -ne $expected[0].Count) { throw ('Wrong cache column count: '+$mode+' '+$phase) }
			for ($r=0; $r -lt $expected.Count; $r++) {
				for ($c=0; $c -lt $expected[$r].Count; $c++) {
					$actual=$summary.Cells.Item($r+1,$c+1).Value2
					$want=$expected[$r][$c]
					if ($want -is [pscustomobject] -and $want.error) {
						if ($summary.Cells.Item($r+1,$c+1).Text -cne $want.error) { throw ('Wrong cache error: '+$mode+' '+$phase) }
					} elseif ($want -is [ValueType] -and $want -isnot [bool]) {
						if ($null -eq $actual) { throw ('Missing numeric cache value: '+$mode+' '+$phase) }
						if ([Math]::Abs([double]$actual-[double]$want) -gt 0.00001) { throw ('Wrong cache value: '+$mode+' '+$phase+' row'+$r+' col'+$c) }
					} elseif ([string]$actual -cne [string]$want) { throw ('Wrong cache label: '+$mode+' '+$phase+' row'+$r+' col'+$c+': '+$actual+' expected: '+$want) }
				}
			}
		}
		Write-Output ($mode+': native Excel accepted every cached and refreshed typed-cache cell')
		$book.Close($false); $book=$null
	}
} finally {
	if ($book) { $book.Close($false) }
	if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
	$env:OOXML_PIVOT_ROOT = $previousRoot; $env:OOXML_PIVOT_OUTPUT = $previousOutput
}
Write-Output ('Temporary acceptance files: '+$taskFolder)
