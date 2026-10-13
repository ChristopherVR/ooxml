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
for (const mode of ['cache-blank', 'cache-numeric', 'cache-boolean', 'cache-error', 'rows', 'rows-none', 'rows-repeat', 'rows-average', 'rows-column', 'rows-sparse', 'rows-blank', 'rows-numeric', 'rows-boolean', 'rows-error', 'rows-mixed']) {
 const wb=createWorkbook({sheets:['Data','Summary']});const session=createEditSession(wb);
 const rows:(string|number|boolean)[][]=[['Category','Region','Amount','Quarter'],['Fruit','East',10,'Q1'],['Fruit','West',20,'Q2'],['Veg','East',5,'Q1'],['Veg','West',15,'Q2'],['Fruit','East',7,'Q2']];
 if(mode==='rows-sparse') rows[3][1]='North';
 for(let r=1;r<rows.length;r++) {
  if(mode.endsWith('-blank') && rows[r][0]==='Fruit') rows[r][0]='';
  if(mode.endsWith('-numeric')) rows[r][0]=rows[r][0]==='Fruit'?2026:2027;
  if(mode.endsWith('-boolean')) rows[r][0]=rows[r][0]==='Fruit';
  if(mode.endsWith('-error') && rows[r][0]==='Fruit') rows[r][0]='=NA()';
 }
 if(mode==='rows-mixed') [10,"'10",true,'Fruit','=NA()'].forEach((value,index)=>rows[index+1][0]=value);
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:D6',rowField:'Category',valueField:'Amount',...(mode.startsWith('cache-')?{columnField:'Region'}:{secondRowField:'Region',subtotals:mode!=='rows-none',repeatLabels:mode==='rows-repeat',aggregate:mode==='rows-average'?'average':'sum',...(mode==='rows-column'?{columnField:'Quarter'}:{})})});
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[1].pivotTables[0]).values));
}
const hierarchy=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-rows-repeat.xlsx'))));
createEditSession(hierarchy).refreshPivotTable(0,'SalesPivot',{repeatLabels:false,subtotals:false});
writeFileSync(join(out,'rows-imported.xlsx'),await saveXlsx(hierarchy));
writeFileSync(join(out,'rows-imported.json'),JSON.stringify(pivotGrid(hierarchy.sheets[0].pivotTables[0]).values));
createEditSession(hierarchy).refreshPivotTable(0,'SalesPivot',{secondRowField:'',columnField:'Region',repeatLabels:true});
writeFileSync(join(out,'cache-from-rows.xlsx'),await saveXlsx(hierarchy));
writeFileSync(join(out,'cache-from-rows.json'),JSON.stringify(pivotGrid(hierarchy.sheets[0].pivotTables[0]).values));
const layouts = ['compact-top','compact-bottom','compact-none','compact-repeat','compact-column','compact-average','compact-numeric-child','cache-compact-single','outline-top','outline-bottom','outline-none','outline-repeat','outline-column','outline-average','cache-outline-single'];
for (const mode of layouts) {
 const wb=createWorkbook({sheets:['Data','Summary']});const session=createEditSession(wb);
 const rows:(string|number)[][]=[['Category','Region','Amount','Quarter'],['Fruit','East',10,'Q1'],['Fruit','West',20,'Q2'],['Veg','East',5,'Q1'],['Veg','West',15,'Q2'],['Fruit','East',7,'Q2']];
 if(mode.endsWith('numeric-child')) for(let r=1;r<rows.length;r++) rows[r][1]=rows[r][1]==='East'?2026:2027;
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:D6',rowField:'Category',valueField:'Amount',layout:mode.includes('compact')?'compact':'outline',subtotalPosition:mode.endsWith('bottom')?'bottom':'top',subtotals:!mode.endsWith('none'),repeatLabels:mode.endsWith('repeat'),aggregate:mode.endsWith('average')?'average':'sum',...(!mode.startsWith('cache-')?{secondRowField:'Region'}:{}),...(mode.endsWith('column')?{columnField:'Quarter'}:{})});
 const grid=pivotGrid(wb.sheets[1].pivotTables[0]);
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(grid.values));
 writeFileSync(join(out,mode+'.indent'),JSON.stringify(grid.indents??grid.values.map(()=>0)));
}
for (const mode of ['cache-values','cache-values-columns','values-rows','values-rows-columns','values-no-subtotals','values-repeat','cache-values-min-max','cache-values-errors','cache-values-two','cache-values-custom']) {
 const wb=createWorkbook({sheets:['Data','Summary']}); const session=createEditSession(wb);
 const rows=[['Category','Region','Amount','Quantity','Year'],['Fruit','East',10,2,2025],['Fruit','West',20,3,2026],['Veg','East',5,4,2025],['Veg','West',15,5,2026],['Fruit','East',7,6,2026]];
 if(mode.endsWith('errors')) rows[1][2]='=NA()';
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 const additionalValues=mode.endsWith('min-max')?[{field:'Amount',aggregate:'min'},{field:'Amount',aggregate:'max'}]:mode.endsWith('two')?[{field:'Quantity',aggregate:'count'}]:mode.endsWith('custom')?[{field:'Quantity',aggregate:'count',caption:'Units'},{field:'Amount',aggregate:'average',caption:'Per item'}]:[{field:'Quantity',aggregate:'count'},{field:'Amount',aggregate:'average'}];
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:E6',rowField:'Category',valueField:'Amount',layout:'tabular',additionalValues,...(!mode.startsWith('cache-')?{secondRowField:'Region'}:{}),...(mode.endsWith('columns')?{columnField:'Year'}:{}),subtotals:mode!=='values-no-subtotals',repeatLabels:mode==='values-repeat'});
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[1].pivotTables[0]).values));
}
for (const profile of ['plain','columns','repeat','columns-repeat','two','min-max','errors','sparse']) {
 const wb=createWorkbook({sheets:['Data','Summary']});const session=createEditSession(wb);
 const rows=[['Category','Region','Amount','Quantity','Year'],['Fruit','East',10,2,2025],['Fruit','West',20,3,2026],['Veg','East',5,4,2025],['Veg','West',15,5,2026],['Fruit','East',7,6,2026]];
 if(profile==='errors') rows[1][2]='=NA()'; if(profile==='sparse') rows[3][4]=2027;
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 const additionalValues=profile==='two'?[{field:'Quantity',aggregate:'count'}]:profile==='min-max'?[{field:'Amount',aggregate:'min'},{field:'Amount',aggregate:'max'}]:[{field:'Quantity',aggregate:'count'},{field:'Amount',aggregate:'average'}];
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:E6',rowField:'Category',valueField:'Amount',layout:'tabular',valuesAxis:'rows',additionalValues,repeatLabels:profile.includes('repeat'),...(profile.includes('columns')||profile==='sparse'?{columnField:'Year'}:{})});
 const mode='row-values-'+profile;
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[1].pivotTables[0]).values));
}
for(const suffix of ['','-columns']){
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-row-values'+suffix+'.xlsx'))));const session=createEditSession(wb);session.setCellInput(1,1,2,'100');session.refreshPivotTable(0,'SalesPivot');
 const mode='row-values-imported'+suffix;writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[0].pivotTables[0]).values));
}
for (const profile of ['to-rows','to-columns','to-single']) {
 const filename=profile==='to-rows'?'excel-pivot-values-columns.xlsx':'excel-pivot-row-values-columns.xlsx';
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__',filename))));const session=createEditSession(wb);
 session.refreshPivotTable(0,'SalesPivot',profile==='to-rows'?{valuesAxis:'rows'}:profile==='to-columns'?{valuesAxis:'columns'}:{valuesAxis:'columns',additionalValues:[]});
 const mode=profile==='to-rows'?'row-values-converted':'cache-values-converted-'+(profile==='to-single'?'single':'columns');
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[0].pivotTables[0]).values));
}
for (const layout of ['compact','outline']) for (const profile of ['one','one-columns','top','bottom','none','top-columns','bottom-columns','none-columns','repeat']) {
 const wb=createWorkbook({sheets:['Data','Summary']}); const session=createEditSession(wb);
 const rows=[['Category','Region','Amount','Quantity','Year'],['Fruit','East',10,2,2025],['Fruit','West',20,3,2026],['Veg','East',5,4,2025],['Veg','West',15,5,2026],['Fruit','East',7,6,2026]];
 rows.forEach((line,r)=>line.forEach((value,c)=>session.setCellInput(0,r,c,String(value))));
 session.createPivotTable(1,{row:0,col:0},{source:'Data!A1:E6',rowField:'Category',valueField:'Amount',layout,subtotalPosition:profile.startsWith('bottom')?'bottom':'top',subtotals:!profile.startsWith('none'),repeatLabels:profile==='repeat',additionalValues:[{field:'Quantity',aggregate:'count'},{field:'Amount',aggregate:'average'}],...(!profile.startsWith('one')?{secondRowField:'Region'}:{}),...(profile.endsWith('columns')?{columnField:'Year'}:{})});
 const mode=(profile.startsWith('one')?'cache-':'')+'values-'+layout+'-'+profile; const grid=pivotGrid(wb.sheets[1].pivotTables[0]);
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb)); writeFileSync(join(out,mode+'.json'),JSON.stringify(grid.values)); writeFileSync(join(out,mode+'.indent'),JSON.stringify(grid.indents??grid.values.map(()=>0)));
}
for (const suffix of ['compact','outline','compact-rows-top-columns','outline-rows-bottom-columns']) {
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-values-'+suffix+'.xlsx')))); const session=createEditSession(wb);
 session.setCellInput(1,1,2,'100'); session.refreshPivotTable(0,'SalesPivot');
 const mode=(suffix.includes('rows')?'':'cache-')+'values-import-'+suffix; const grid=pivotGrid(wb.sheets[0].pivotTables[0]);
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb)); writeFileSync(join(out,mode+'.json'),JSON.stringify(grid.values)); writeFileSync(join(out,mode+'.indent'),JSON.stringify(grid.indents??grid.values.map(()=>0)));
}
for (const flag of ['visible','omitted']) {
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-visible-values.xlsx'))));
 const session=createEditSession(wb); const pivot=wb.sheets[0].pivotTables[0];
 if(flag==='omitted') pivot.sourceXml=pivot.sourceXml.replace('hideValuesRow="0"','');
 session.refreshPivotTable(0,pivot.name,{additionalValues:[{field:'Amount',aggregate:'count'}]});
 const mode='cache-values-'+flag+'-two';
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[0].pivotTables[0]).values));
}
for (const [mode,suffix] of [['cache-values-imported',''],['values-imported','-rows-columns'],['cache-values-single','']]) {
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-values'+suffix+'.xlsx'))));
 const session=createEditSession(wb); session.setCellInput(1,1,2,'100'); session.refreshPivotTable(0,'SalesPivot',mode.endsWith('single')?{additionalValues:[],layout:'compact'}:{});
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(pivotGrid(wb.sheets[0].pivotTables[0]).values));
}
for (const suffix of ['compact-top','compact-bottom','compact-bottom-no-subtotals','outline-top','outline-bottom','outline-top-repeat']) {
 const wb=await loadXlsx(new Uint8Array(readFileSync(join(root,'src/core/xlsx/__fixtures__/excel-pivot-'+suffix+'.xlsx'))));
 const session=createEditSession(wb);session.setCellInput(1,1,2,'100');session.refreshPivotTable(0,'SalesPivot');
 const mode='import-'+suffix; const grid=pivotGrid(wb.sheets[0].pivotTables[0]);
 writeFileSync(join(out,mode+'.xlsx'),await saveXlsx(wb));
 writeFileSync(join(out,mode+'.json'),JSON.stringify(grid.values));
 writeFileSync(join(out,mode+'.indent'),JSON.stringify(grid.indents));
}
'@ | Set-Content -LiteralPath $taskScript
	& bun run $taskScript
	if ($LASTEXITCODE -ne 0) { throw 'PivotTable generation failed.' }
	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false; $excel.DisplayAlerts = $false
	Write-Output ('Native Excel version: '+$excel.Version+', build: '+$excel.Build)
	Write-Output ('Native Excel file version: '+[Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $excel.Path 'EXCEL.EXE')).FileVersion)
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
	foreach ($expectedFile in Get-ChildItem -LiteralPath $taskFolder -Filter '*.json') {
		$mode=$expectedFile.BaseName
		$book = $excel.Workbooks.Open((Join-Path $taskFolder ($mode+'.xlsx')))
		$summary = $book.Worksheets.Item('Summary')
		$pivot = $summary.PivotTables().Item(1)
		if ($mode.Contains('values') -and $pivot.DataFields().Count -ne $(if($mode.EndsWith('single')){1}elseif($mode.EndsWith('two')){2}else{3})) { throw ('Wrong PivotTable measure count: '+$mode) }
		$rowFieldCount = if ($mode.StartsWith('cache-')) { 1 } else { 2 }
		if ($pivot.RowFields().Count -ne $rowFieldCount) { throw ('Wrong PivotTable row field count: '+$mode) }
		$expected = Get-Content -Raw -LiteralPath (Join-Path $taskFolder ($mode+'.json')) | ConvertFrom-Json
		foreach ($phase in @('cached','refreshed')) {
			if ($phase -eq 'refreshed') { [void]$pivot.RefreshTable() }
			if ($pivot.TableRange2.Rows.Count -ne $expected.Count) { throw ('Wrong hierarchy row count: '+$mode+' '+$phase) }
			if ($pivot.TableRange2.Columns.Count -ne $expected[0].Count) { throw ('Wrong hierarchy column count: '+$mode+' '+$phase) }
			for ($r=0; $r -lt $expected.Count; $r++) {
				for ($c=0; $c -lt $expected[$r].Count; $c++) {
					$actual=$summary.Cells.Item($r+1,$c+1).Value2
					$want=$expected[$r][$c]
					if ($want -is [pscustomobject] -and $want.error) {
						if ($summary.Cells.Item($r+1,$c+1).Text -cne $want.error) { throw ('Wrong hierarchy error: '+$mode+' '+$phase) }
					} elseif ($want -is [ValueType] -and $want -isnot [bool]) {
						if ($null -eq $actual) { throw ('Missing numeric hierarchy value: '+$mode+' '+$phase) }
						if ($actual -isnot [ValueType] -or $actual -is [bool]) { throw ('Wrong numeric hierarchy type: '+$mode+' '+$phase) }
						if ([Math]::Abs([double]$actual-[double]$want) -gt 0.00001) { throw ('Wrong hierarchy value: '+$mode+' '+$phase+' row'+$r+' col'+$c) }
					} elseif ([string]$actual -cne [string]$want -or ($want -is [string] -and $actual -isnot [string])) { throw ('Wrong hierarchy label: '+$mode+' '+$phase+' row'+$r+' col'+$c+': '+$actual+' expected: '+$want) }
				}
			}
			$indentFile=Join-Path $taskFolder ($mode+'.indent')
			if (Test-Path -LiteralPath $indentFile) {
				$indents=Get-Content -Raw -LiteralPath $indentFile | ConvertFrom-Json
				for ($r=0; $r -lt $indents.Count; $r++) {
					if ($summary.Cells.Item($r+1,1).IndentLevel -ne $indents[$r]) { throw ('Wrong report indentation: '+$mode+' '+$phase+' row'+$r) }
					if ($indents[$r] -eq 1 -and $summary.Cells.Item($r+1,1).HorizontalAlignment -ne -4131) { throw ('Wrong compact child alignment: '+$mode+' '+$phase) }
				}
			}
		}
		Write-Output ($mode+': native Excel accepted every cached and refreshed hierarchy cell')
		$book.Close($false); $book=$null
	}
	$book = $excel.Workbooks.Open((Join-Path $taskRoot 'src/core/xlsx/__fixtures__/excel-pivot-values-compact-rows-top-columns.xlsx'))
	$summary = $book.Worksheets.Item('Summary'); $pivot = $summary.PivotTables().Item(1)
	$pivot.DataPivotField.Name = 'Metrics'
	if ($summary.Range('B1').Text -ne 'Column Labels') { throw 'Native compact header exposed a custom Values caption.' }
	$pivot.RowAxisLayout(2)
	if ($summary.Range('C1').Text -ne 'Year' -or $summary.Range('D1').Text -ne 'Metrics') { throw 'Native outline custom Values caption was not surfaced.' }
	Write-Output 'custom Values caption: hidden in compact, visible after switching to outline (unsupported preserved profile)'
	$book.Close($false); $book=$null
} finally {
	if ($book) { $book.Close($false) }
	if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
	$env:OOXML_PIVOT_ROOT = $previousRoot; $env:OOXML_PIVOT_OUTPUT = $previousOutput
}
Write-Output ('Temporary acceptance files: '+$taskFolder)
