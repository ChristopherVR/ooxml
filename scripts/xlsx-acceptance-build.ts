// Builds the workbooks `xlsx-excel-acceptance.ps1` opens in Excel. Run with bun:
//   bun scripts/xlsx-acceptance-build.ts <output folder>
// Every file is written by `saveXlsx`: new workbooks built from `createWorkbook()`, the
// committed fixtures saved untouched, and fixtures saved after edits that force the writer to
// regenerate comments, tables and drawings.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateSync } from 'node:zlib';
import { cellError, type CellStyle, type Workbook } from '../src/xlsx/model.js';
import { putCell } from '../src/xlsx/cells.js';
import { internStyle } from '../src/xlsx/styles.js';
import { createWorkbook, createWorksheet, defaultCellStyle } from '../src/xlsx/workbook.js';
import { loadXlsx } from '../src/xlsx/read/index.js';
import { saveXlsx } from '../src/xlsx/write/index.js';
import { type CellCheck, sessionBuilds } from './xlsx-acceptance-session.js';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/xlsx-acceptance-build.ts <output folder>');
mkdirSync(out, { recursive: true });
const fixtures = join(import.meta.dir, '..', 'src', 'xlsx', '__fixtures__');
const range = (a: [number, number], b: [number, number]) => ({
	start: { row: a[0], col: a[1] },
	end: { row: b[0], col: b[1] },
});
const style = (wb: Workbook, patch: Partial<CellStyle>) =>
	internStyle(wb, { ...defaultCellStyle(), ...patch });

function basic(): Workbook {
	const wb = createWorkbook({ sheets: ['Values', 'Second'] });
	const sheet = wb.sheets[0]!;
	const bold = style(wb, {
		font: { ...defaultCellStyle().font, bold: true, color: { rgb: 'FFFFFFFF' } },
		fill: { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } },
	});
	const money = style(wb, { numFmt: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' });
	const date = style(wb, { numFmt: 'yyyy-mm-dd' });
	const boxed = style(wb, {
		border: {
			left: { style: 'thin' },
			right: { style: 'thin' },
			top: { style: 'double', color: { rgb: 'FFFF0000' } },
			bottom: { style: 'medium' },
		},
		alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
	});
	const gradient = style(wb, {
		fill: {
			type: 'gradient',
			gradient: 'linear',
			degree: 90,
			stops: [
				{ position: 0, color: { rgb: 'FFFFFFFF' } },
				{ position: 1, color: { theme: 5 } },
			],
		},
	});
	['Name', 'Amount', 'Date', 'Flag', 'Error', 'Formula'].forEach((text, col) =>
		putCell(sheet, 0, col, { value: text, styleId: bold }),
	);
	for (let row = 1; row <= 20; row++) {
		putCell(sheet, row, 0, { value: `Item ${row}` });
		putCell(sheet, row, 1, { value: row * 12.5 - 40, styleId: money });
		putCell(sheet, row, 2, { value: 45000 + row, styleId: date });
		putCell(sheet, row, 3, { value: row % 2 === 0 });
		putCell(sheet, row, 4, {
			value: row === 3 ? cellError('#N/A') : null,
			...(row === 3 ? {} : { styleId: boxed }),
		});
		putCell(sheet, row, 5, { value: row * 12.5 - 40, formula: `B${row + 1}*1`, styleId: gradient });
	}
	putCell(sheet, 22, 0, { value: 'Line 1\r\nLine 2 with _x0041_ literal', styleId: boxed });
	putCell(sheet, 23, 0, {
		value: 'Rich',
		richText: [
			{ text: 'Ri', font: { bold: true, color: { rgb: 'FFFF0000' } } },
			{ text: 'ch', font: { italic: true, name: 'Arial', size: 14 } },
		],
	});
	putCell(sheet, 24, 1, { value: 1, formula: 'XLOOKUP("Item 2",A2:A21,B2:B21)' });
	putCell(sheet, 25, 1, {
		value: 2,
		formula: 'SUM(B2:B21*F2:F21)',
		arrayRange: range([25, 1], [25, 1]),
	});
	putCell(sheet, 26, 1, { value: 'big', formula: 'IF(B26>1,"big","small")' });
	sheet.merges.push(range([28, 0], [29, 2]));
	putCell(sheet, 28, 0, { value: 'Merged', styleId: boxed });
	sheet.columns.push(
		{ min: 0, max: 0, width: 20, customWidth: true },
		{ min: 6, max: 7, hidden: true, width: 9 },
	);
	sheet.rowInfo.set(22, { height: 30, customHeight: true });
	sheet.rowInfo.set(30, { hidden: true });
	sheet.view.freeze = { rows: 1, cols: 1 };
	sheet.view.zoom = 110;
	sheet.tabColor = { rgb: 'FF00B050' };
	wb.sheets[1]!.state = 'hidden';
	putCell(wb.sheets[1]!, 0, 0, { value: 3, formula: 'Values!B2+1' });
	wb.properties = { title: 'Acceptance basic', creator: 'xlsx acceptance' };
	return wb;
}

function features(): Workbook {
	const wb = createWorkbook({ sheets: ['Data', 'Notes'] });
	const sheet = wb.sheets[0]!;
	['Region', 'Q1', 'Q2'].forEach((text, col) => putCell(sheet, 0, col, { value: text }));
	['North', 'South', 'East', 'West'].forEach((name, i) => {
		putCell(sheet, i + 1, 0, { value: name });
		putCell(sheet, i + 1, 1, { value: 10 + i * 3 });
		putCell(sheet, i + 1, 2, { value: 20 - i * 2 });
	});
	putCell(sheet, 5, 0, { value: 'Total' });
	putCell(sheet, 5, 2, { value: 68, formula: 'SUBTOTAL(109,Sales[Q2])' });
	sheet.tables.push({
		id: 1,
		name: 'Sales',
		displayName: 'Sales',
		range: range([0, 0], [5, 2]),
		headerRow: true,
		totalsRow: true,
		columns: [
			{ name: 'Region', totalsRowLabel: 'Total' },
			{ name: 'Q1' },
			{ name: 'Q2', totalsRowFunction: 'sum' },
		],
		styleName: 'TableStyleMedium9',
		showRowStripes: true,
	});
	const red = {
		font: { color: { rgb: 'FF9C0006' } },
		fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { rgb: 'FFFFC7CE' } },
	};
	sheet.conditionalFormats.push(
		{
			ranges: [range([1, 1], [4, 1])],
			rules: [
				{ type: 'cellIs', operator: 'greaterThan', formulas: ['12'], style: red, priority: 1 },
			],
		},
		{
			ranges: [range([1, 2], [4, 2])],
			rules: [
				{
					type: 'colorScale',
					thresholds: [{ type: 'min' }, { type: 'max' }],
					colors: [{ rgb: 'FFF8696B' }, { rgb: 'FF63BE7B' }],
					priority: 2,
				},
			],
		},
		{
			ranges: [range([1, 1], [4, 1])],
			rules: [
				{
					type: 'dataBar',
					min: { type: 'min' },
					max: { type: 'max' },
					color: { rgb: 'FF638EC6' },
					priority: 3,
				},
			],
		},
		{
			ranges: [range([1, 2], [4, 2])],
			rules: [
				{
					type: 'iconSet',
					iconSet: '3Arrows',
					thresholds: [
						{ type: 'percent', value: '0' },
						{ type: 'percent', value: '33' },
						{ type: 'percent', value: '67' },
					],
					priority: 4,
				},
			],
		},
		{
			ranges: [range([1, 0], [4, 0])],
			rules: [
				{ type: 'containsText', text: 'th', style: red, priority: 5 },
				{
					type: 'expression',
					formula: 'MOD(ROW(),2)=0',
					style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: 0.8 } } },
					priority: 6,
				},
				{ type: 'top10', rank: 2, style: { font: { bold: true } }, priority: 7 },
				{ type: 'aboveAverage', below: true, style: { font: { italic: true } }, priority: 8 },
				{ type: 'duplicateValues', style: { font: { strike: true } }, priority: 9 },
				{
					type: 'containsBlanks',
					style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFEEEEEE' } } },
					priority: 10,
				},
			],
		},
	);
	sheet.dataValidations.push(
		{
			ranges: [range([1, 4], [10, 4])],
			type: 'list',
			formula1: '"Yes,No,Maybe"',
			allowBlank: true,
			showDropDown: true,
			showErrorMessage: true,
			error: 'Pick one',
			errorTitle: 'Invalid',
		},
		{
			ranges: [range([1, 5], [10, 5])],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			showInputMessage: true,
			prompt: '1 to 10',
		},
		{
			ranges: [range([1, 6], [10, 6])],
			type: 'date',
			operator: 'greaterThan',
			formula1: 'DATE(2020,1,1)',
		},
	);
	putCell(sheet, 8, 0, { value: 'Example' });
	putCell(sheet, 9, 0, { value: 'Go to notes' });
	sheet.hyperlinks.push(
		{ range: range([8, 0], [8, 0]), target: 'https://example.com/?a=1&b=2', tooltip: 'Example' },
		{ range: range([9, 0], [9, 0]), location: 'Notes!A1', display: 'Go to notes' },
	);
	sheet.comments.push(
		{ address: { row: 1, col: 0 }, author: 'Tester', text: 'Plain note\nwith two lines' },
		{
			address: { row: 2, col: 0 },
			author: 'Tester',
			text: 'Thread start',
			replies: [{ author: 'Reviewer', text: 'A reply' }],
		},
	);
	sheet.pageSetup = {
		orientation: 'landscape',
		paperSize: 9,
		fitToWidth: 1,
		fitToHeight: 0,
		printArea: range([0, 0], [10, 6]),
		header: '&CQuarterly',
		footer: '&RPage &P',
		margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
	};
	wb.sheets[1]!.protection = { sheet: true, allow: ['formatCells', 'sort'] };
	wb.sheets[1]!.autoFilter = {
		range: range([0, 0], [3, 1]),
		columns: [{ offset: 0, values: ['a', 'b'] }],
	};
	['Key', 'Value', 'a', 1, 'b', 2, 'c', 3].forEach((value, i) =>
		putCell(wb.sheets[1]!, Math.floor(i / 2), i % 2, { value }),
	);
	wb.definedNames.push(
		{ name: 'Rate', formula: 'Data!$B$2' },
		{ name: 'LocalName', formula: 'Notes!$A$1', localSheet: 1 },
	);
	wb.activeSheet = 0;
	return wb;
}

/** A 2x2 solid red PNG built by hand (zlib from node, CRC-32 inline). */
function tinyPng(): Uint8Array {
	const crcTable = Array.from({ length: 256 }, (_, n) => {
		let c = n;
		for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		return c >>> 0;
	});
	const crc = (bytes: Buffer) => {
		let c = 0xffffffff;
		for (const b of bytes) c = (crcTable[(c ^ b) & 0xff] ?? 0) ^ (c >>> 8);
		return (c ^ 0xffffffff) >>> 0;
	};
	const chunk = (type: string, data: Buffer) => {
		const head = Buffer.alloc(4);
		head.writeUInt32BE(data.length);
		const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
		const tail = Buffer.alloc(4);
		tail.writeUInt32BE(crc(body));
		return Buffer.concat([head, body, tail]);
	};
	const header = Buffer.alloc(13);
	header.writeUInt32BE(2, 0);
	header.writeUInt32BE(2, 4);
	header.set([8, 2, 0, 0, 0], 8);
	const rows = Buffer.from([0, 255, 0, 0, 255, 0, 0, 0, 255, 0, 0, 255, 0, 0]);
	return new Uint8Array(
		Buffer.concat([
			Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
			chunk('IHDR', header),
			chunk('IDAT', deflateSync(rows)),
			chunk('IEND', Buffer.alloc(0)),
		]),
	);
}
const PNG = tinyPng();

function charts(): Workbook {
	const wb = createWorkbook({ sheets: ['Chart data'] });
	const sheet = wb.sheets[0]!;
	['Month', 'Sales', 'Cost'].forEach((text, col) => putCell(sheet, 0, col, { value: text }));
	['Jan', 'Feb', 'Mar', 'Apr'].forEach((m, i) => {
		putCell(sheet, i + 1, 0, { value: m });
		putCell(sheet, i + 1, 1, { value: 100 + i * 10 });
		putCell(sheet, i + 1, 2, { value: 60 + i * 5 });
	});
	const series = (col: string, name: string, values: number[]) => ({
		name,
		nameRef: `'Chart data'!$${col}$1`,
		categoriesRef: "'Chart data'!$A$2:$A$5",
		valuesRef: `'Chart data'!$${col}$2:$${col}$5`,
		categories: ['Jan', 'Feb', 'Mar', 'Apr'],
		values,
	});
	const anchor = (row: number, col: number) => ({
		from: { row, col, rowOffset: 0, colOffset: 0 },
		to: { row: row + 14, col: col + 7, rowOffset: 0, colOffset: 0 },
	});
	sheet.drawings.push(
		{
			kind: 'chart',
			anchor: anchor(0, 4),
			chartType: 'column',
			grouping: 'clustered',
			title: 'Sales and cost',
			showLegend: true,
			legendPosition: 'b',
			series: [series('B', 'Sales', [100, 110, 120, 130]), series('C', 'Cost', [60, 65, 70, 75])],
		},
		{
			kind: 'chart',
			anchor: anchor(16, 4),
			chartType: 'line',
			showLegend: true,
			series: [series('B', 'Sales', [100, 110, 120, 130])],
		},
		{
			kind: 'chart',
			anchor: anchor(32, 4),
			chartType: 'pie',
			title: 'Share',
			showLegend: true,
			series: [series('C', 'Cost', [60, 65, 70, 75])],
		},
		{
			kind: 'chart',
			anchor: anchor(0, 13),
			chartType: 'bar',
			grouping: 'stacked',
			showLegend: false,
			series: [series('B', 'Sales', [100, 110, 120, 130]), series('C', 'Cost', [60, 65, 70, 75])],
		},
		{
			kind: 'chart',
			anchor: anchor(16, 13),
			chartType: 'scatter',
			showLegend: true,
			series: [
				{
					...series('C', 'Cost', [60, 65, 70, 75]),
					categoriesRef: "'Chart data'!$B$2:$B$5",
					categories: [100, 110, 120, 130],
				},
			],
		},
		{
			kind: 'image',
			anchor: {
				from: { row: 7, col: 0, rowOffset: 0, colOffset: 0 },
				ext: { cx: 952500, cy: 952500 },
			},
			partName: 'xl/media/image1.png',
			contentType: 'image/png',
			name: 'Logo',
		},
	);
	wb.source = { parts: new Map([['xl/media/image1.png', PNG]]) };
	return wb;
}

async function edited(name: string): Promise<Workbook> {
	const wb = await loadXlsx(readFileSync(join(fixtures, name)));
	for (const sheet of wb.sheets) {
		putCell(sheet, 40, 0, { value: 'edited by saveXlsx' });
		if (sheet.comments.length)
			sheet.comments.push({ address: { row: 40, col: 1 }, author: 'Editor', text: 'Added note' });
		for (const object of sheet.drawings)
			object.anchor = {
				...object.anchor,
				from: { ...object.anchor.from, row: object.anchor.from.row + 1 },
				...(object.anchor.to ? { to: { ...object.anchor.to, row: object.anchor.to.row + 1 } } : {}),
			};
		for (const table of sheet.tables) table.showColumnStripes = !table.showColumnStripes;
		sheet.view.zoom = 120;
		if (sheet.pageSetup) sheet.pageSetup = { ...sheet.pageSetup, orientation: 'portrait' };
	}
	wb.sheets.push(createWorksheet('Added', Math.max(...wb.sheets.map((s) => s.sheetId)) + 1));
	wb.theme = { ...wb.theme, colors: wb.theme.colors.map((c, i) => (i === 4 ? 'C00000' : c)) };
	return wb;
}

/** Mirrors the edit module's `duplicateSheet`: no part names on the copy, its tables or charts. */
async function duplicated(name: string): Promise<Workbook> {
	const wb = await loadXlsx(readFileSync(join(fixtures, name)));
	const count = wb.sheets.length;
	for (let i = 0; i < count; i++) {
		const copy = structuredClone(wb.sheets[i]!);
		copy.name = `${copy.name} (2)`.slice(0, 31);
		copy.sheetId = 100 + i;
		delete copy.partName;
		copy.tables = copy.tables.map((t, n) => {
			const next = {
				...t,
				id: 100 + i * 10 + n,
				name: `${t.name}_${i + 2}`,
				displayName: `${t.name}_${i + 2}`,
			};
			delete next.partName;
			return next;
		});
		for (const drawing of copy.drawings) if (drawing.kind === 'chart') delete drawing.partName;
		wb.sheets.push(copy);
	}
	return wb;
}

const builds: [string, () => Promise<Workbook> | Workbook][] = [
	['new-empty.xlsx', () => createWorkbook()],
	['new-basic.xlsx', basic],
	['new-features.xlsx', features],
	['new-charts.xlsx', charts],
];
const FIXTURES = [
	'excel-features.xlsx',
	'excel-1904.xlsx',
	'openpyxl-features.xlsx',
	'openpyxl-styles.xlsx',
	'openpyxl-1904.xlsx',
];
const checks: Record<string, CellCheck[]> = {};
for (const [file, build, cellChecks] of sessionBuilds(fixtures, FIXTURES, PNG)) {
	builds.push([
		file,
		async () => {
			const workbook = await build();
			checks[file] = cellChecks();
			return workbook;
		},
	]);
}
for (const fixture of FIXTURES) {
	builds.push([
		`roundtrip-${fixture}`,
		async () => loadXlsx(readFileSync(join(fixtures, fixture))),
	]);
	builds.push([`edited-${fixture}`, () => edited(fixture)]);
	builds.push([`duplicated-${fixture}`, () => duplicated(fixture)]);
}
/** What Excel should report for each sheet; the PowerShell script compares it over COM. */
const manifest: Record<
	string,
	{
		name: string;
		tables: number;
		charts: number;
		hyperlinks: number;
		comments: number;
		merges: number;
	}[]
> = {};
for (const [file, build] of builds) {
	const workbook = await build();
	let bytes = await saveXlsx(workbook);
	// A second generation checks that our own output survives another load and save.
	const reloaded = await loadXlsx(bytes);
	bytes = await saveXlsx(reloaded);
	writeFileSync(join(out, file), bytes);
	manifest[file] = reloaded.sheets.map((sheet) => ({
		name: sheet.name,
		tables: sheet.tables.length,
		charts: sheet.drawings.filter((d) => d.kind === 'chart').length,
		hyperlinks: sheet.hyperlinks.length,
		comments: sheet.comments.length,
		merges: sheet.merges.length,
	}));
	console.log(`${file}\t${bytes.length}`);
}
writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, '\t'));
writeFileSync(join(out, 'checks.json'), JSON.stringify(checks, null, '\t'));
