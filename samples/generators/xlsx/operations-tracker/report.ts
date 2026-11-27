/**
 * The Depot report: three printable sections over the log (by depot, by category and priority,
 * by technician), set up for landscape pages with repeated titles and a break between sections.
 */
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import type { Book } from '../lib/book';
import { L, rangeOf } from '../lib/sheet';
import { CATEGORIES, DEPOTS, PRIORITIES, STATUSES } from './data';
import { FMT, LINE, P, SHEETS, TOTAL, header, section } from './style';

const S = SHEETS.report;
const DEPOT_HEAD = 6;
const CATEGORY_HEAD = DEPOT_HEAD + DEPOTS.length + 5;
const TECH_HEAD = CATEGORY_HEAD + CATEGORIES.length + 5;
const TECHNICIANS = DEPOTS.flatMap((depot) =>
	depot.technicians.map((name) => ({ name, depot: depot.name })),
);
const LAST = TECH_HEAD + TECHNICIANS.length + 1;

/** Column sums for a total row; `ratio` columns are recomputed instead of summed. */
function totals(
	book: Book,
	row: number,
	first: number,
	last: number,
	columns: string[],
	formats: Record<string, string>,
): void {
	const sheet = book.sheet(S);
	sheet.paint(`B${row}:L${row}`, TOTAL);
	sheet.set(`B${row}`, 'Total', { ...TOTAL, indent: 1 });
	for (const col of columns)
		sheet.formula(`${col}${row}`, `SUM(${col}${first}:${col}${last})`, {
			...TOTAL,
			fmt: formats[col] ?? FMT.count,
		});
}

function byDepot(book: Book): void {
	const sheet = book.sheet(S);
	const first = DEPOT_HEAD + 1;
	const last = DEPOT_HEAD + DEPOTS.length;
	section(sheet, DEPOT_HEAD - 1, 'B', 'L', 'By depot');
	header(sheet, DEPOT_HEAD, [
		['B', 'Depot', 'left'],
		...STATUSES.map((status, i): [string, string] => [L(2 + i), status]),
		['G', 'Total'],
		['H', 'Overdue'],
		['I', 'Critical open'],
		['J', 'Closed on time'],
		['K', 'Labour hours'],
		['L', 'Parts cost'],
	]);
	DEPOTS.forEach((depot, n) => {
		const r = first + n;
		sheet.paint(`B${r}:L${r}`, LINE);
		sheet.set(`B${r}`, depot.name, { ...LINE, bold: true, indent: 1 });
		STATUSES.forEach((_, i) =>
			sheet.formula(
				`${L(2 + i)}${r}`,
				`COUNTIFS(WO_Depot,$B${r},WO_Status,${L(2 + i)}$${DEPOT_HEAD})`,
				{ ...LINE, fmt: FMT.count },
			),
		);
		sheet.formula(`G${r}`, `SUM(C${r}:F${r})`, { ...LINE, fmt: FMT.count, bold: true });
		sheet.formula(`H${r}`, `COUNTIFS(WO_Depot,$B${r},WO_Overdue,">0")`, {
			...LINE,
			fmt: FMT.count,
		});
		sheet.formula(
			`I${r}`,
			`COUNTIFS(WO_Depot,$B${r},WO_Priority,"Critical",WO_Status,"<>Closed")`,
			{ ...LINE, fmt: FMT.count },
		);
		// Closed no later than the due date, as a share of the closed work.
		sheet.formula(
			`J${r}`,
			`SUMPRODUCT((WO_Depot=$B${r})*(WO_Status="Closed")*(WO_Closed<=WO_Due))/F${r}`,
			{ ...LINE, fmt: FMT.percent },
		);
		sheet.formula(`K${r}`, `SUMIFS(WO_Hours,WO_Depot,$B${r})`, { ...LINE, fmt: FMT.count });
		sheet.formula(`L${r}`, `SUMIFS(WO_Cost,WO_Depot,$B${r})`, { ...LINE, fmt: FMT.money });
	});
	totals(book, last + 1, first, last, ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L'], {
		L: FMT.money,
	});
	sheet.formula(
		`J${last + 1}`,
		`SUMPRODUCT((WO_Status="Closed")*(WO_Closed<=WO_Due))/F${last + 1}`,
		{ ...TOTAL, fmt: FMT.percent },
	);
	sheet.conditional(`H${first}:H${last}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: { rgb: 'FFEDB7AC' },
	});
	sheet.conditional(`J${first}:J${last}`, {
		type: 'iconSet',
		iconSet: '3Symbols',
		thresholds: [
			{ type: 'percent', value: '0' },
			{ type: 'num', value: '0.8' },
			{ type: 'num', value: '0.84' },
		],
	});
}

function byCategory(book: Book): void {
	const sheet = book.sheet(S);
	const first = CATEGORY_HEAD + 1;
	const last = CATEGORY_HEAD + CATEGORIES.length;
	section(sheet, CATEGORY_HEAD - 1, 'B', 'L', 'By category and priority');
	header(sheet, CATEGORY_HEAD, [
		['B', 'Category', 'left'],
		...PRIORITIES.map((priority, i): [string, string] => [L(2 + i), priority]),
		['G', 'Total'],
		['H', 'Open'],
		['I', 'Share of work'],
		['J', 'Hours per job'],
		['K', 'Labour hours'],
		['L', 'Parts cost'],
	]);
	CATEGORIES.forEach((category, n) => {
		const r = first + n;
		sheet.paint(`B${r}:L${r}`, LINE);
		sheet.set(`B${r}`, category, { ...LINE, bold: true, indent: 1 });
		PRIORITIES.forEach((_, i) =>
			sheet.formula(
				`${L(2 + i)}${r}`,
				`COUNTIFS(WO_Category,$B${r},WO_Priority,${L(2 + i)}$${CATEGORY_HEAD})`,
				{ ...LINE, fmt: FMT.count },
			),
		);
		sheet.formula(`G${r}`, `SUM(C${r}:F${r})`, { ...LINE, fmt: FMT.count, bold: true });
		sheet.formula(`H${r}`, `COUNTIFS(WO_Category,$B${r},WO_Status,"<>Closed")`, {
			...LINE,
			fmt: FMT.count,
		});
		sheet.formula(`I${r}`, `G${r}/G$${last + 1}`, { ...LINE, fmt: FMT.percent });
		sheet.formula(`J${r}`, `K${r}/G${r}`, { ...LINE, fmt: FMT.hours });
		sheet.formula(`K${r}`, `SUMIFS(WO_Hours,WO_Category,$B${r})`, { ...LINE, fmt: FMT.count });
		sheet.formula(`L${r}`, `SUMIFS(WO_Cost,WO_Category,$B${r})`, { ...LINE, fmt: FMT.money });
	});
	totals(book, last + 1, first, last, ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L'], {
		I: FMT.percent,
		L: FMT.money,
	});
	sheet.formula(`J${last + 1}`, `K${last + 1}/G${last + 1}`, { ...TOTAL, fmt: FMT.hours });
	sheet.conditional(`C${first}:F${last}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'percentile', value: '60' }, { type: 'max' }],
		colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FFDCE5F2' }, { rgb: 'FF9DB6DB' }],
	});
}

function byTechnician(book: Book): void {
	const sheet = book.sheet(S);
	const first = TECH_HEAD + 1;
	const last = TECH_HEAD + TECHNICIANS.length;
	section(sheet, TECH_HEAD - 1, 'B', 'L', 'By technician');
	header(sheet, TECH_HEAD, [
		['B', 'Technician', 'left'],
		['C', 'Depot', 'left'],
		['E', 'Assigned'],
		['F', 'Closed'],
		['G', 'Open'],
		['H', 'Overdue'],
		['I', 'Closed on time'],
		['J', 'Days to close'],
		['K', 'Labour hours'],
		['L', 'Parts cost'],
	]);
	sheet.paint(`D${TECH_HEAD}`, { fill: P.navy });
	TECHNICIANS.forEach((person, n) => {
		const r = first + n;
		sheet.paint(`B${r}:L${r}`, LINE);
		sheet.set(`B${r}`, person.name, { ...LINE, bold: true, indent: 1 });
		sheet.set(`C${r}`, person.depot, { ...LINE, indent: 1 });
		sheet.formula(`E${r}`, `COUNTIFS(WO_Tech,$B${r})`, { ...LINE, fmt: FMT.count });
		sheet.formula(`F${r}`, `COUNTIFS(WO_Tech,$B${r},WO_Status,"Closed")`, {
			...LINE,
			fmt: FMT.count,
		});
		sheet.formula(`G${r}`, `E${r}-F${r}`, { ...LINE, fmt: FMT.count });
		sheet.formula(`H${r}`, `COUNTIFS(WO_Tech,$B${r},WO_Overdue,">0")`, { ...LINE, fmt: FMT.count });
		sheet.formula(
			`I${r}`,
			`SUMPRODUCT((WO_Tech=$B${r})*(WO_Status="Closed")*(WO_Closed<=WO_Due))/F${r}`,
			{ ...LINE, fmt: FMT.percent },
		);
		sheet.formula(`J${r}`, `AVERAGEIFS(WO_Age,WO_Tech,$B${r},WO_Status,"Closed")`, {
			...LINE,
			fmt: FMT.days,
		});
		sheet.formula(`K${r}`, `SUMIFS(WO_Hours,WO_Tech,$B${r})`, { ...LINE, fmt: FMT.count });
		sheet.formula(`L${r}`, `SUMIFS(WO_Cost,WO_Tech,$B${r})`, { ...LINE, fmt: FMT.money });
	});
	totals(book, last + 1, first, last, ['E', 'F', 'G', 'H', 'K', 'L'], { L: FMT.money });
	sheet.formula(
		`I${last + 1}`,
		`SUMPRODUCT((WO_Status="Closed")*(WO_Closed<=WO_Due))/F${last + 1}`,
		{ ...TOTAL, fmt: FMT.percent },
	);
	sheet.formula(`J${last + 1}`, 'AVERAGEIFS(WO_Age,WO_Status,"Closed")', {
		...TOTAL,
		fmt: FMT.days,
	});
	sheet.conditional(`G${first}:G${last}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: { rgb: 'FFB9C7DD' },
	});
	sheet.conditional(`I${first}:I${last}`, {
		type: 'top10',
		rank: 3,
		style: { font: { bold: true, color: { rgb: `FF${P.teal}` } } },
	});
	sheet.conditional(`I${first}:I${last}`, {
		type: 'top10',
		rank: 3,
		bottom: true,
		style: { font: { bold: true, color: { rgb: `FF${P.signal}` } } },
	});
}

export function buildReport(book: Book): void {
	const sheet = book.sheet(S);
	sheet.present(P.amber);
	sheet.widths({
		A: 2,
		B: 25,
		C: 11,
		D: 13,
		E: 17,
		F: 11,
		G: 11,
		H: 11,
		I: 14,
		J: 15,
		K: 13,
		L: 13,
	});
	sheet.heights({ 1: 9, 2: 27, 3: 16 });
	sheet.set('B2', 'Depot report', { font: 'display', size: 16, bold: true, color: P.navy });
	sheet.formula(
		'B3',
		'"Maintenance work orders as of "&TEXT(AsOf,"d mmmm yyyy")&". Every count is a formula over the Work orders sheet."',
		{ color: P.muted, size: 9, italic: true },
	);
	byDepot(book);
	byCategory(book);
	byTechnician(book);
	sheet.ws.pageSetup = {
		orientation: 'landscape',
		paperSize: 9,
		scale: 88,
		printArea: rangeOf(`B1:L${LAST}`),
		margins: { left: 0.5, right: 0.5, top: 0.8, bottom: 0.7, header: 0.35, footer: 0.35 },
		header: '&L&BVale Orrin Logistics&B&RDepot report',
		footer: '&LPrinted from the operations tracker&CPage &P of &N&RAs of 15 January 2026',
	};
	sheet.ws.printOptions = { horizontalCentered: true };
}

/** Print titles and the page breaks before the second and third sections. */
export function printReport(book: Book, session: EditSession): void {
	const index = book.sheet(S).index;
	session.setPrintTitles(index, { rows: { start: 0, end: 2 } });
	session.setPageBreak(index, 'row', CATEGORY_HEAD - 2);
	session.setPageBreak(index, 'row', TECH_HEAD - 2);
}
