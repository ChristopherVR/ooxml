# Run from the worktree with Windows PowerShell and desktop Excel installed.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root ('node_modules/.cache/pivot-context-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output -Force | Out-Null
$generator = Join-Path $output 'generate.ts'
$excel = $null
$owned = $false
$runningExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

Add-Type -Name ExcelWindow -Namespace PivotAcceptance -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
'@

function Assert-ReportCells($Pivot, $Expected) {
	$values = $Pivot.TableRange2.Value2
	if ($values.GetLength(0) -ne $Expected.Count -or $values.GetLength(1) -ne $Expected[0].Count) {
		throw 'Excel returned a different report size.'
	}
	for ($row = 0; $row -lt $Expected.Count; $row++) {
		for ($col = 0; $col -lt $Expected[$row].Count; $col++) {
			if ($Expected[$row][$col] -cne $values[($row + 1), ($col + 1)]) {
				throw "Excel returned a different value at report offset ($row, $col)."
			}
		}
	}
}

try {
	@'
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createEditSession, loadXlsx, pivotGrid, saveXlsx, type PivotTablePatch } from '../../../src/core/xlsx/index';

const cases = {
	design: {
		name: 'NewSales',
		layout: 'outline',
		repeatLabels: true,
		subtotals: false,
		display: {
			colGrandTotals: false,
			showHeaders: false,
			showDrill: false,
			blankRows: true,
			style: { name: 'PivotStyleMedium4', showRowStripes: true, showColStripes: true },
		},
	},
	dark: { display: { style: { name: 'PivotStyleDark28' } } },
	unstyled: { display: { style: { name: '' } } },
	collapsed: { display: { collapsed: true } },
	caption: { condition: { kind: 'captionEqual', value: 'Veg' } },
	'caption-pattern': { condition: { kind: 'captionEqual', value: 'f?u*' } },
	'caption-escape': { condition: { kind: 'captionEqual', value: '~~~*~?' } },
	empty: { condition: { kind: 'captionEqual', value: 'Missing' } },
	greater: { condition: { kind: 'valueGreaterThan', measure: 0, value: 20 } },
	less: { condition: { kind: 'valueLessThan', measure: 0, value: 30 } },
	calculated: { calculatedField: { name: 'DoubleAmount', field: 'Amount', factor: 2 } },
	moved: { name: 'NewSales', destination: { row: 3, col: 5 }, source: 'Data!A1:C5' },
	'totals-rows': { display: { rowGrandTotals: true, colGrandTotals: false } },
	'totals-columns': { display: { rowGrandTotals: false, colGrandTotals: true } },
} satisfies Record<string, PivotTablePatch>;

for (const [name, patch] of Object.entries(cases)) {
	const suffix = name.startsWith('totals-') ? '' : '-rows';
	const fixture = readFileSync(join(import.meta.dir, '../../../src/core/xlsx/__fixtures__/excel-pivot' + suffix + '.xlsx'));
	const workbook = await loadXlsx(fixture);
	const edit = createEditSession(workbook);
	if (name === 'caption-escape') edit.setCellInput(1, 1, 0, '~*?');
	edit.refreshPivotTable(0, 'SalesPivot', patch);
	const pivot = workbook.sheets[0]!.pivotTables![0]!;
	writeFileSync(join(import.meta.dir, name + '.xlsx'), await saveXlsx(workbook));
	writeFileSync(join(import.meta.dir, name + '.json'), JSON.stringify(pivotGrid(pivot).values));
}
'@ | Set-Content -LiteralPath $generator -Encoding UTF8
	& bun $generator
	if ($LASTEXITCODE -ne 0) { throw 'Workbook generation failed.' }

	$excel = New-Object -ComObject Excel.Application
	[uint32]$excelProcess = 0
	[void][PivotAcceptance.ExcelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcess)
	if (!$excelProcess -or $runningExcel -contains $excelProcess) {
		throw 'Excel automation did not start a separate instance.'
	}
	$owned = $true
	$excel.DisplayAlerts = $false
	$excel.AutomationSecurity = 3
	Write-Output ("Excel " + $excel.Version + " build " + $excel.Build)
	$expectedFilters = @{ caption = 15; 'caption-pattern' = 15; 'caption-escape' = 15; empty = 15; greater = 9; less = 11 }
	foreach ($file in Get-ChildItem -LiteralPath $output -Filter '*.xlsx') {
		$case = $file.BaseName
		$expected = Get-Content -LiteralPath (Join-Path $output ($case + '.json')) -Raw | ConvertFrom-Json
		$book = $excel.Workbooks.Open($file.FullName, 0, $true)
		try {
			$pivot = $book.Worksheets.Item('Summary').PivotTables().Item(1)
			$field = $pivot.PivotFields('Category')
			Assert-ReportCells $pivot $expected
			if ($case -like 'totals-*' -and (
				$pivot.RowGrand -ne ($case -eq 'totals-rows') -or
				$pivot.ColumnGrand -ne ($case -eq 'totals-columns')
			)) {
				throw 'Excel did not read the saved grand-total settings.'
			}
			if ($case -eq 'design') {
				if ($pivot.Name -ne 'NewSales' -or $pivot.TableStyle2.Name -ne 'PivotStyleMedium4' -or
					!$pivot.ShowTableStyleRowStripes -or !$pivot.ShowTableStyleColumnStripes -or
					$field.Subtotals(1) -or $pivot.ShowDrillIndicators) {
					throw 'Excel did not read the saved design options.'
				}
			}
			if ($case -eq 'dark' -and $pivot.TableStyle2.Name -ne 'PivotStyleDark28') {
				throw 'Excel did not read the saved dark style.'
			}
			if ($case -eq 'unstyled' -and $pivot.TableStyle2.Name) {
				throw 'Excel did not clear the style.'
			}
			if ($expectedFilters.ContainsKey($case) -and
				$field.PivotFilters().Item(1).FilterType -ne $expectedFilters[$case]) {
				throw 'Excel did not read the saved filter.'
			}
			if ($case -eq 'calculated' -and $pivot.CalculatedFields().Item(1).Name -ne 'DoubleAmount') {
				throw 'Excel did not read the saved calculated field.'
			}
			if ($case -eq 'moved' -and ($pivot.Name -ne 'NewSales' -or $pivot.TableRange2.Address() -ne '$F$4:$H$11')) {
				throw 'Excel did not read the saved name or destination.'
			}
			if (!$pivot.RefreshTable()) { throw 'RefreshTable returned false.' }
			Assert-ReportCells $pivot $expected
			$pivot.RowAxisLayout(2)
			$field.Subtotals = @($false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false,$false)
			$pivot.ShowTableStyleRowStripes = $true
			Write-Output "$case passed open, refresh, cell comparison and design edits"
		} finally {
			$book.Close($false)
		}
	}
} finally {
	if ($excel) {
		if ($owned) { $excel.Quit() }
		[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
	}
	Get-ChildItem -LiteralPath $output -File | Remove-Item
	Remove-Item -LiteralPath $output
}
