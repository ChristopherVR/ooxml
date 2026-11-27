/**
 * The Analysis sheet: every aggregate the dashboard draws, computed from the Orders table with
 * structured references so the figures follow the data.
 */
import { ymdToSerial } from '../../../../src/core/xlsx/formula/date-serial';
import type { Book, Look } from '../lib/book';
import { L } from '../lib/sheet';
import { CATEGORIES, CHANNELS, type Order, TERRITORIES, YEAR, revenueOf } from './data';
import { FMT, FOOTER, HEAD, P, SHEETS, section, title } from './style';

const S = SHEETS.analysis;

/** Where each block lives, for the dashboard and the charts. */
export const AT = {
	monthHead: 6,
	monthFirst: 7,
	monthLast: 18,
	monthTotal: 19,
	categoryHead: 23,
	categoryFirst: 24,
	categoryLast: 29,
	channelFirst: 24,
	channelLast: 26,
	regionHead: 34,
	regionFirst: 35,
	regionLast: 39,
	countHead: 44,
	countFirst: 45,
	countLast: 49,
} as const;

const total: Look = { bold: true, top: P.ink, bottom: ['double', P.ink] };
const months = Array.from({ length: 12 }, (_, i) => ymdToSerial(YEAR, i + 1, 1, false)!);

function head(book: Book, row: number, from: number, names: string[]): void {
	const sheet = book.sheet(S);
	names.forEach((name, i) =>
		sheet.set(`${L(from + i)}${row}`, name, {
			...HEAD,
			h: i === 0 ? 'left' : 'right',
			...(i === 0 ? { indent: 1 } : {}),
		}),
	);
	sheet.heights({ [row]: 20 });
}

function monthly(book: Book): void {
	const sheet = book.sheet(S);
	section(sheet, 5, 'B', 'M', 'Monthly performance');
	head(book, AT.monthHead, 1, [
		'Month',
		'Revenue',
		'Cost',
		'Margin',
		'Margin %',
		'Orders',
		'Units',
		'Avg order',
		'Cumulative',
		...CHANNELS,
	]);
	const cells: [string, (r: number) => string, string][] = [
		['C', (r) => `SUMIFS(Orders[Revenue],Orders[Month],$B${r})`, FMT.money],
		['D', (r) => `SUMIFS(Orders[Cost],Orders[Month],$B${r})`, FMT.money],
		['E', (r) => `C${r}-D${r}`, FMT.money],
		['F', (r) => `E${r}/C${r}`, FMT.percent],
		['G', (r) => `COUNTIFS(Orders[Month],$B${r})`, FMT.count],
		['H', (r) => `SUMIFS(Orders[Units],Orders[Month],$B${r})`, FMT.count],
		['I', (r) => `C${r}/G${r}`, FMT.money],
		['J', (r) => `SUM(C$${AT.monthFirst}:C${r})`, FMT.money],
		...CHANNELS.map((_, i): [string, (r: number) => string, string] => [
			L(10 + i),
			(r) =>
				`SUMIFS(Orders[Revenue],Orders[Month],$B${r},Orders[Channel],${L(10 + i)}$${AT.monthHead})`,
			FMT.money,
		]),
	];
	months.forEach((serial, i) => {
		const r = AT.monthFirst + i;
		sheet.set(`B${r}`, serial, { fmt: FMT.monthYear, h: 'left', indent: 1, bottom: P.mist });
		for (const [col, formula, fmt] of cells)
			sheet.formula(`${col}${r}`, formula(r), { fmt, bottom: P.mist });
	});
	const t = AT.monthTotal;
	const span = (col: string) => `${col}${AT.monthFirst}:${col}${AT.monthLast}`;
	sheet.set(`B${t}`, 'Year', { ...total, indent: 1 });
	for (const col of ['C', 'D', 'E', 'G', 'H', 'K', 'L', 'M'])
		sheet.formula(`${col}${t}`, `SUM(${span(col)})`, {
			...total,
			fmt: col === 'G' || col === 'H' ? FMT.count : FMT.money,
		});
	sheet.formula(`F${t}`, `E${t}/C${t}`, { ...total, fmt: FMT.percent });
	sheet.formula(`I${t}`, `C${t}/G${t}`, { ...total, fmt: FMT.money });
	sheet.paint(`J${t}`, total);
	// The best and the weakest month of the year.
	sheet.conditional(
		`C${AT.monthFirst}:C${AT.monthLast}`,
		{ type: 'top10', rank: 1, style: { font: { bold: true, color: { rgb: `FF${P.pine}` } } } },
		{
			type: 'top10',
			rank: 1,
			bottom: true,
			style: { font: { bold: true, color: { rgb: `FF${P.rust}` } } },
		},
	);
	sheet.conditional(`F${AT.monthFirst}:F${AT.monthLast}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'max' }],
		colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FFBFD9D1' }],
	});
}

/** Categories in descending revenue order, so the bar chart reads as a ranking. */
function rankedCategories(list: Order[]): string[] {
	const revenue = new Map<string, number>();
	for (const order of list)
		revenue.set(
			order.product.category,
			(revenue.get(order.product.category) ?? 0) + revenueOf(order),
		);
	return [...CATEGORIES].sort((a, b) => (revenue.get(b) ?? 0) - (revenue.get(a) ?? 0));
}

function breakdowns(book: Book, list: Order[]): void {
	const sheet = book.sheet(S);
	section(sheet, 22, 'B', 'G', 'By category');
	head(book, AT.categoryHead, 1, ['Category', 'Revenue', 'Margin', 'Margin %', 'Units', 'Share']);
	rankedCategories(list).forEach((category, i) => {
		const r = AT.categoryFirst + i;
		const line: Look = { bottom: P.mist };
		sheet.set(`B${r}`, category, { ...line, indent: 1 });
		sheet.formula(`C${r}`, `SUMIFS(Orders[Revenue],Orders[Category],$B${r})`, {
			...line,
			fmt: FMT.money,
		});
		sheet.formula(`D${r}`, `SUMIFS(Orders[Margin],Orders[Category],$B${r})`, {
			...line,
			fmt: FMT.money,
		});
		sheet.formula(`E${r}`, `D${r}/C${r}`, { ...line, fmt: FMT.percent });
		sheet.formula(`F${r}`, `SUMIFS(Orders[Units],Orders[Category],$B${r})`, {
			...line,
			fmt: FMT.count,
		});
		sheet.formula(`G${r}`, `C${r}/C$${AT.categoryLast + 1}`, { ...line, fmt: FMT.percent });
	});
	const t = AT.categoryLast + 1;
	sheet.set(`B${t}`, 'All categories', { ...total, indent: 1 });
	for (const col of ['C', 'D', 'F'])
		sheet.formula(`${col}${t}`, `SUM(${col}${AT.categoryFirst}:${col}${AT.categoryLast})`, {
			...total,
			fmt: col === 'F' ? FMT.count : FMT.money,
		});
	sheet.formula(`E${t}`, `D${t}/C${t}`, { ...total, fmt: FMT.percent });
	sheet.formula(`G${t}`, `SUM(G${AT.categoryFirst}:G${AT.categoryLast})`, {
		...total,
		fmt: FMT.percent,
	});
	sheet.conditional(`C${AT.categoryFirst}:C${AT.categoryLast}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: { rgb: 'FFA9CDC2' },
	});

	section(sheet, 22, 'I', 'L', 'By channel');
	head(book, AT.categoryHead, 8, ['Channel', 'Revenue', 'Orders', 'Share']);
	CHANNELS.forEach((channel, i) => {
		const r = AT.channelFirst + i;
		const line: Look = { bottom: P.mist };
		sheet.set(`I${r}`, channel, { ...line, indent: 1 });
		sheet.formula(`J${r}`, `SUMIFS(Orders[Revenue],Orders[Channel],$I${r})`, {
			...line,
			fmt: FMT.money,
		});
		sheet.formula(`K${r}`, `COUNTIFS(Orders[Channel],$I${r})`, { ...line, fmt: FMT.count });
		sheet.formula(`L${r}`, `J${r}/J$${AT.channelLast + 1}`, { ...line, fmt: FMT.percent });
	});
	const c = AT.channelLast + 1;
	sheet.set(`I${c}`, 'All channels', { ...total, indent: 1 });
	sheet.formula(`J${c}`, `SUM(J${AT.channelFirst}:J${AT.channelLast})`, {
		...total,
		fmt: FMT.money,
	});
	sheet.formula(`K${c}`, `SUM(K${AT.channelFirst}:K${AT.channelLast})`, {
		...total,
		fmt: FMT.count,
	});
	sheet.formula(`L${c}`, `SUM(L${AT.channelFirst}:L${AT.channelLast})`, {
		...total,
		fmt: FMT.percent,
	});
}

/** A region by month grid: revenue, or the number of orders. */
function regionGrid(
	book: Book,
	top: number,
	heading: string,
	cell: (r: number, col: string) => string,
	fmt: string,
): void {
	const sheet = book.sheet(S);
	section(sheet, top - 1, 'B', 'O', heading);
	sheet.set(`B${top}`, 'Region', { ...HEAD, indent: 1 });
	months.forEach((serial, i) =>
		sheet.set(`${L(2 + i)}${top}`, serial, { ...HEAD, h: 'right', fmt: FMT.month }),
	);
	sheet.set(`O${top}`, 'Year', { ...HEAD, h: 'right' });
	sheet.heights({ [top]: 20 });
	TERRITORIES.forEach((territory, n) => {
		const r = top + 1 + n;
		sheet.set(`B${r}`, territory.region, { indent: 1, bottom: P.mist });
		months.forEach((_, i) =>
			sheet.formula(`${L(2 + i)}${r}`, cell(r, L(2 + i)), { fmt, bottom: P.mist }),
		);
		sheet.formula(`O${r}`, `SUM(C${r}:N${r})`, { fmt, bold: true, bottom: P.mist });
	});
	const t = top + 1 + TERRITORIES.length;
	sheet.set(`B${t}`, 'All regions', { ...total, indent: 1 });
	for (let i = 0; i <= 12; i++)
		sheet.formula(`${L(2 + i)}${t}`, `SUM(${L(2 + i)}${top + 1}:${L(2 + i)}${t - 1})`, {
			...total,
			fmt,
		});
}

export function buildAnalysis(book: Book, list: Order[]): void {
	const sheet = book.sheet(S);
	title(
		sheet,
		'Analysis',
		'Aggregates of the Orders table. Every figure is a formula with structured references; the dashboard and its charts read from here.',
		P.pine,
	);
	sheet.page(FOOTER, { tall: 1 });
	sheet.widths({ A: 2, B: 22, 'C:O': 12.5 });
	monthly(book);
	breakdowns(book, list);
	regionGrid(
		book,
		AT.regionHead,
		'Revenue by region and month',
		(r, col) =>
			`SUMIFS(Orders[Revenue],Orders[Region],$B${r},Orders[Month],${col}$${AT.regionHead})`,
		FMT.money,
	);
	regionGrid(
		book,
		AT.countHead,
		'Orders by region and month',
		(r, col) => `COUNTIFS(Orders[Region],$B${r},Orders[Month],${col}$${AT.countHead})`,
		FMT.count,
	);
	sheet.conditional(`C${AT.regionFirst}:N${AT.regionLast}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
		colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FFDCEBE5' }, { rgb: 'FF8FBFB0' }],
	});
	book.wb.definedNames.push(
		{ name: 'Months', formula: `${S}!$B$${AT.monthFirst}:$B$${AT.monthLast}` },
		{ name: 'MonthlyRevenue', formula: `${S}!$C$${AT.monthFirst}:$C$${AT.monthLast}` },
		{ name: 'MonthlyOrders', formula: `${S}!$G$${AT.monthFirst}:$G$${AT.monthLast}` },
	);
}
