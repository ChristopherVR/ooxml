// Acceptance workbooks produced through `createEditSession` (dynamic arrays, structural edits,
// sheet operations with tables) plus the fixtures edited in a session. Used by
// `xlsx-acceptance-build.ts`; each build also lists cell checks the PowerShell script runs
// against Excel over COM (values, formulas and dynamic-array spills).
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { formatAddress } from '../src/xlsx/address.js';
import { getCell } from '../src/xlsx/cells.js';
import { createEditSession, type EditSession } from '../src/xlsx/edit/index.js';
import type { CellValue, Workbook } from '../src/xlsx/model.js';
import { isCellError } from '../src/xlsx/model.js';
import { loadXlsx } from '../src/xlsx/read/index.js';
import { createWorkbook } from '../src/xlsx/workbook.js';

/** One cell Excel must agree on after opening (and recalculating) the saved file. */
export interface CellCheck {
	sheet: string;
	cell: string;
	/** Expected `Value2` (numbers compared with a tolerance, errors as their text). */
	value?: string | number | boolean;
	/** Expected `HasSpill` and the spill size. */
	spill?: { rows: number; cols: number };
	/** Expected `Formula2` (with the leading `=`). */
	formula?: string;
}

export type SessionBuild = [
	name: string,
	build: () => Promise<Workbook>,
	checks: () => CellCheck[],
];

const range = (r1: number, c1: number, r2 = r1, c2 = c1) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

/** The value (and formula) check for a cell as our engine computed it. */
function check(wb: Workbook, sheet: number, row: number, col: number): CellCheck {
	const ws = wb.sheets[sheet]!;
	const cell = getCell(ws, row, col);
	const value: CellValue = cell?.value ?? null;
	const out: CellCheck = { sheet: ws.name, cell: formatAddress({ row, col }) };
	if (isCellError(value)) out.value = value.error;
	else if (value !== null) out.value = value;
	if (cell?.formula) out.formula = `=${cell.formula}`;
	return out;
}

function dynamicArrays(): { wb: Workbook; s: EditSession } {
	const wb = createWorkbook({ sheets: ['Dynamic'] });
	const s = createEditSession(wb);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Name', 'Score'],
		['Ann', 31],
		['Bob', 12],
		['Cy', 25],
		['Di', 40],
	]);
	s.setCellInput(0, 0, 3, '=SEQUENCE(4,2,10,5)');
	s.setCellInput(0, 0, 6, '=FILTER(A2:A5,B2:B5>20)');
	s.setCellInput(0, 0, 7, '=SORT(B2:B5,1,-1)');
	s.setCellInput(0, 0, 8, '=UNIQUE({1;2;2;3})');
	s.setCellInput(0, 6, 0, '=SUM(D1#)');
	s.setCellInput(0, 6, 1, '=SUM(B2:B5*(B2:B5>20))');
	s.setCellInput(0, 6, 2, '=B2:B5*2');
	s.setCellValue(0, 13, 4, 'blocker');
	s.setCellInput(0, 12, 4, '=SEQUENCE(3)');
	// Typing into a spill range blocks it; undo restores the spill.
	s.setCellInput(0, 2, 3, 'x');
	s.undo();
	return { wb, s };
}

function structure(): { wb: Workbook; s: EditSession } {
	const wb = createWorkbook({ sheets: ['Ledger', 'Summary'] });
	const s = createEditSession(wb);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Item', 'Qty', 'Price'],
		['a', 2, 3],
		['b', 4, 5],
		['c', 6, 7],
	]);
	for (let r = 1; r <= 3; r++) s.setCellInput(0, r, 3, `=B${r + 1}*C${r + 1}`);
	s.setCellInput(0, 4, 3, '=SUM(D2:D4)');
	s.setCellInput(1, 0, 0, '=Ledger!D5');
	s.setCellInput(1, 1, 0, '=SUM(Ledger!B:B)');
	s.insertRows(0, 2, 2);
	s.setRangeValues(0, { row: 2, col: 0 }, [
		['x', 1, 10],
		['y', 2, 20],
	]);
	s.setCellInput(0, 2, 3, '=B3*C3');
	s.setCellInput(0, 3, 3, '=B4*C4');
	s.deleteRows(0, 1, 1);
	s.insertColumns(0, 0, 1);
	s.deleteColumns(0, 3, 1);
	s.paste(0, { row: 10, col: 6 }, s.cut(0, range(0, 1, 5, 3)));
	return { wb, s };
}

function sheets(): { wb: Workbook; s: EditSession } {
	const wb = createWorkbook({ sheets: ['Data', 'Links', 'Scratch'] });
	const s = createEditSession(wb);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Region', 'Amount'],
		['N', 10],
		['S', 20],
		['E', 30],
	]);
	const table = s.createTable(0, range(0, 0, 3, 1), true);
	s.setCellInput(0, 0, 3, `=SUM(${table.name}[Amount])`);
	s.setCellInput(1, 0, 0, '=Data!B2*2');
	s.setCellInput(1, 1, 0, `=COUNTA(${table.name}[Region])`);
	s.setCellInput(1, 2, 0, '=Scratch!A1');
	s.setDefinedName({ name: 'Rate', formula: 'Data!$B$3' });
	s.setCellInput(1, 3, 0, '=Rate+1');
	s.setCellValue(1, 4, 0, 'To data');
	s.setHyperlink(1, range(4, 0), { location: 'Data!A1', display: 'To data' });
	s.renameSheet(0, 'Sales Data');
	const copy = s.duplicateSheet(0);
	s.setCellValue(copy, 1, 1, 100);
	s.deleteSheet(2 + (copy <= 2 ? 1 : 0));
	s.moveSheet(copy, 0);
	return { wb, s };
}

/** The newer session commands: tables, CF, charts, pictures, styles, page, outline, protection. */
function commands(png: Uint8Array): { wb: Workbook; s: EditSession } {
	const wb = createWorkbook({ sheets: ['Commands', 'Locked'] });
	const s = createEditSession(wb);
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Item', 'Qty', 'Date'],
		['a', 3, 46000],
		['b', 5, 46001],
		['a', 3, 46000],
		['c', 7, 46002],
	]);
	s.removeDuplicates(0, range(0, 0, 4, 2), [], true);
	const table = s.createTable(0, range(0, 0, 3, 2), true);
	s.setCellInput(0, 0, 5, `=SUM(${table.name}[Qty])`);
	s.updateTable(0, table.name, { name: 'Stock', totalsRow: true, showColumnStripes: true });
	s.addConditionalFormat(0, {
		ranges: [range(1, 2, 3, 2)],
		rules: [
			{ type: 'timePeriod', timePeriod: 'lastMonth', style: { font: { bold: true } }, priority: 1 },
		],
	});
	s.addConditionalFormat(0, {
		ranges: [range(1, 1, 3, 1)],
		rules: [
			{
				type: 'cellIs',
				operator: 'greaterThan',
				formulas: ['4'],
				style: { font: { italic: true } },
				priority: 2,
			},
		],
	});
	s.moveConditionalRule(0, 1, 0, 'up');
	s.applyCellStyle(0, [range(0, 5)], 'Good');
	s.applyCellStyle(0, [range(1, 5)], 'Heading 1');
	s.addChart(0, {
		anchor: {
			from: { row: 8, col: 0, rowOffset: 0, colOffset: 0 },
			to: { row: 20, col: 6, rowOffset: 0, colOffset: 0 },
		},
		chartType: 'column',
		title: 'Qty',
		showLegend: true,
		series: [
			{
				name: 'Qty',
				valuesRef: 'Commands!$B$2:$B$4',
				categoriesRef: 'Commands!$A$2:$A$4',
				categories: [],
				values: [],
			},
		],
	});
	s.updateChart(0, 0, { title: 'Quantity by item', legendPosition: 'b' });
	s.addImage(
		0,
		png,
		'image/png',
		{ from: { row: 0, col: 8, rowOffset: 0, colOffset: 0 }, ext: { cx: 952500, cy: 952500 } },
		'Logo',
	);
	s.setDrawingAnchor(0, 1, {
		from: { row: 2, col: 8, rowOffset: 0, colOffset: 0 },
		ext: { cx: 952500, cy: 952500 },
	});
	s.setPageSetup(0, { orientation: 'landscape', fitToWidth: 1, fitToHeight: 0 });
	s.setPrintOptions(0, { gridLines: true, horizontalCentered: true });
	s.groupRows(0, 1, 3);
	s.groupColumns(0, 1, 2);
	s.setDefaultColumnWidth(0, 12);
	s.setSheetView(0, { showFormulas: true });
	s.setCellValue(1, 0, 0, 'secret');
	s.setSheetProtection(1, { sheet: true, allow: ['formatCells'] }, 'pw');
	s.setWorkbookProtection(true, 'book');
	s.setCalcMode('manual');
	return { wb, s };
}

/** The fixture edits of the integration test, through a session. */
async function sessionEdited(fixtures: string, name: string): Promise<Workbook> {
	const wb = await loadXlsx(readFileSync(join(fixtures, name)));
	const s = createEditSession(wb);
	const first = wb.sheets.findIndex((sheet) => sheet.state === 'visible');
	s.insertRows(first, 1, 2);
	s.insertColumns(first, 0, 1);
	s.setCellInput(first, 0, 0, '=SEQUENCE(2,2)');
	s.setCellInput(first, 3, 0, '=SUM(A1:B2)');
	s.deleteRows(first, 5, 1);
	s.paste(first, { row: 40, col: 10 }, s.cut(first, range(3, 0)));
	s.renameSheet(first, `${wb.sheets[first]!.name.slice(0, 20)} R`);
	const copy = s.duplicateSheet(first);
	s.addSheet();
	s.setCellInput(wb.sheets.length - 1, 0, 0, `='${wb.sheets[copy]!.name}'!A1*2`);
	if (wb.sheets.length > 3) s.deleteSheet(wb.sheets.length - 2);
	return wb;
}

export function sessionBuilds(
	fixtures: string,
	names: readonly string[],
	png: Uint8Array,
): SessionBuild[] {
	const dyn = dynamicArrays();
	const cmd = commands(png);
	const str = structure();
	const sh = sheets();
	const builds: SessionBuild[] = [
		[
			'session-dynamic.xlsx',
			async () => dyn.wb,
			() => [
				{ sheet: 'Dynamic', cell: 'D1', spill: { rows: 4, cols: 2 }, value: 10 },
				{ sheet: 'Dynamic', cell: 'G1', spill: { rows: 3, cols: 1 }, value: 'Ann' },
				{ sheet: 'Dynamic', cell: 'H1', spill: { rows: 4, cols: 1 }, value: 40 },
				{ sheet: 'Dynamic', cell: 'I1', spill: { rows: 3, cols: 1 }, value: 1 },
				{ sheet: 'Dynamic', cell: 'C7', spill: { rows: 4, cols: 1 }, value: 62 },
				check(dyn.wb, 0, 6, 0),
				check(dyn.wb, 0, 6, 1),
				{ sheet: 'Dynamic', cell: 'E13', value: '#SPILL!' },
				{ sheet: 'Dynamic', cell: 'D3', value: 30 },
			],
		],
		[
			'session-structure.xlsx',
			async () => str.wb,
			() => [
				...[10, 11, 12, 13, 14, 15].flatMap((r) => [6, 7, 8].map((c) => check(str.wb, 0, r, c))),
				check(str.wb, 1, 0, 0),
				check(str.wb, 1, 1, 0),
			],
		],
		[
			'session-sheets.xlsx',
			async () => sh.wb,
			() => {
				const links = sh.wb.sheets.findIndex((x) => x.name === 'Links');
				const data = sh.wb.sheets.findIndex((x) => x.name === 'Sales Data');
				return [
					...[0, 1, 2, 3].map((r) => check(sh.wb, links, r, 0)),
					check(sh.wb, data, 0, 3),
					check(sh.wb, 0, 0, 3),
				];
			},
		],
	];
	builds.push([
		'session-commands.xlsx',
		async () => cmd.wb,
		() => [
			{ sheet: 'Commands', cell: 'F1', value: 15, formula: '=SUM(Stock[Qty])' },
			{ sheet: 'Commands', cell: 'C5', value: 138003 },
		],
	]);
	for (const name of names)
		builds.push([`session-${name}`, () => sessionEdited(fixtures, name), () => []]);
	return builds;
}
