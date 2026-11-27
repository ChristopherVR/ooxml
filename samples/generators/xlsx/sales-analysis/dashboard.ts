/** The Dashboard: headline tiles, four charts and the region table with sparklines. */
import { type Book, type Look, rgb } from '../lib/book';
import { type Tile, tiles as drawTiles } from '../lib/tiles';
import { AT } from './analysis';
import { TERRITORIES } from './data';
import { FMT, FOOTER, HEAD, P, SHEETS, section } from './style';

const S = SHEETS.dashboard;
const DATA = SHEETS.analysis;
/** The region table: its header row and its first and last region rows. */
const REGION_HEAD = 48;
export const REGION_FIRST = REGION_HEAD + 1;
export const REGION_LAST = REGION_HEAD + TERRITORIES.length;

function tiles(book: Book): void {
	const revenue = 'SUM(Orders[Revenue])';
	const count = 'ROWS(Orders[Order ID])';
	const specs: Tile[] = [
		{
			first: 'B',
			last: 'D',
			label: 'Revenue',
			value: revenue,
			fmt: FMT.thousands,
			note: '"Best month: "&TEXT(INDEX(Months,MATCH(MAX(MonthlyRevenue),MonthlyRevenue,0)),"mmmm")',
		},
		{
			first: 'E',
			last: 'G',
			label: 'Order lines',
			value: count,
			fmt: FMT.count,
			note: `"Average order "&TEXT(${revenue}/${count},"$#,##0")`,
		},
		{
			first: 'H',
			last: 'J',
			label: 'Gross margin',
			value: `SUM(Orders[Margin])/${revenue}`,
			fmt: FMT.percent,
			note: '"Margin earned: "&TEXT(SUM(Orders[Margin]),"$#,##0")',
		},
		{
			first: 'K',
			last: 'M',
			label: 'Against annual target',
			value: `${revenue}/SUM(Regions[Annual target])`,
			fmt: FMT.percent,
			note: `"Target "&TEXT(SUM(Regions[Annual target]),"$#,##0")&" across "&ROWS(Regions[Region])&" regions"`,
		},
	];
	drawTiles(book.sheet(S), 5, specs, { fill: P.paper, accent: P.pine, muted: P.muted });
}

function regions(book: Book): void {
	const sheet = book.sheet(S);
	section(sheet, REGION_HEAD - 1, 'B', 'M', 'Regions');
	sheet.paint(`B${REGION_HEAD}:M${REGION_HEAD}`, HEAD);
	const heads: [string, string, NonNullable<Look['h']>][] = [
		['B', 'Region', 'left'],
		['D', 'Regional manager', 'left'],
		['F', 'Revenue', 'right'],
		['G', 'Share', 'right'],
		['H', 'Vs target', 'right'],
		['I', 'Orders', 'right'],
		['J', 'Avg order', 'right'],
		['K', 'Revenue trend', 'center'],
		['L', 'Orders trend', 'center'],
		['M', 'Best month', 'right'],
	];
	for (const [col, text, h] of heads)
		sheet.set(`${col}${REGION_HEAD}`, text, {
			...HEAD,
			h,
			// The two sparkline columns are narrow: their headers take two lines.
			...(col === 'K' || col === 'L' ? { wrap: true } : {}),
			...(col === 'B' ? { indent: 1 } : {}),
		});
	sheet.heights({ [REGION_HEAD]: 30 });
	const totalRow = REGION_LAST + 1;
	TERRITORIES.forEach((_, n) => {
		const r = REGION_FIRST + n;
		const source = AT.regionFirst + n;
		const counts = AT.countFirst + n;
		const line: Look = { bottom: P.mist, v: 'center' };
		sheet.paint(`B${r}:M${r}`, line);
		sheet.formula(`B${r}`, `${DATA}!B${source}`, { ...line, bold: true, indent: 1 });
		sheet.formula(`D${r}`, `XLOOKUP(B${r},Regions[Region],Regions[Regional manager])`, line);
		sheet.formula(`F${r}`, `${DATA}!O${source}`, { ...line, fmt: FMT.money });
		sheet.formula(`G${r}`, `F${r}/F$${totalRow}`, { ...line, fmt: FMT.percent });
		sheet.formula(`H${r}`, `F${r}/XLOOKUP(B${r},Regions[Region],Regions[Annual target])-1`, {
			...line,
			fmt: FMT.delta,
		});
		sheet.formula(`I${r}`, `${DATA}!O${counts}`, { ...line, fmt: FMT.count });
		sheet.formula(`J${r}`, `F${r}/I${r}`, { ...line, fmt: FMT.money });
		sheet.formula(
			`M${r}`,
			`INDEX(${DATA}!$C$${AT.regionHead}:$N$${AT.regionHead},MATCH(MAX(${DATA}!C${source}:N${source}),${DATA}!C${source}:N${source},0))`,
			{ ...line, fmt: 'mmmm', color: P.muted },
		);
		sheet.heights({ [r]: 24 });
	});
	const total: Look = { bold: true, top: P.ink, bottom: ['double', P.ink], v: 'center' };
	sheet.paint(`B${totalRow}:M${totalRow}`, total);
	sheet.set(`B${totalRow}`, 'All regions', { ...total, indent: 1 });
	sheet.formula(`F${totalRow}`, `SUM(F${REGION_FIRST}:F${REGION_LAST})`, {
		...total,
		fmt: FMT.money,
	});
	sheet.formula(`G${totalRow}`, `SUM(G${REGION_FIRST}:G${REGION_LAST})`, {
		...total,
		fmt: FMT.percent,
	});
	sheet.formula(`H${totalRow}`, `F${totalRow}/SUM(Regions[Annual target])-1`, {
		...total,
		fmt: FMT.delta,
	});
	sheet.formula(`I${totalRow}`, `SUM(I${REGION_FIRST}:I${REGION_LAST})`, {
		...total,
		fmt: FMT.count,
	});
	sheet.formula(`J${totalRow}`, `F${totalRow}/I${totalRow}`, { ...total, fmt: FMT.money });
	sheet.heights({ [totalRow]: 20 });
	sheet.conditional(`F${REGION_FIRST}:F${REGION_LAST}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: rgb('A9CDC2'),
	});
	sheet.conditional(`H${REGION_FIRST}:H${REGION_LAST}`, {
		type: 'iconSet',
		iconSet: '3Symbols2',
		thresholds: [
			{ type: 'percent', value: '0' },
			{ type: 'num', value: '-0.05' },
			{ type: 'num', value: '0' },
		],
	});
}

export function buildDashboard(book: Book, lines: number): void {
	const sheet = book.sheet(S);
	sheet.present(P.pine).page(FOOTER, { portrait: true, tall: 1 });
	sheet.widths({ A: 2, 'B:M': 11.5 });
	sheet.heights({ 1: 9, 2: 30, 3: 17, 4: 9 });
	sheet.set('B2', 'Kestrelline Outfitters', {
		font: 'display',
		size: 20,
		bold: true,
		color: P.pine,
	});
	sheet.set(
		'B3',
		`Sales analysis for calendar year 2025: ${lines} order lines across five regions and three channels.`,
		{ color: P.muted },
	);
	sheet.set('M3', 'Explore the data', { color: P.lake, underline: true, h: 'right' });
	sheet.linkTo('M3', `${SHEETS.explorer}!A1`, 'Open the PivotTable, slicers and timeline');
	tiles(book);
	section(sheet, 9, 'B', 'M', 'Revenue through the year');
	section(sheet, 28, 'B', 'M', 'Where the revenue comes from');
	regions(book);
	sheet.set(
		`B${REGION_LAST + 3}`,
		'Kestrelline Outfitters is a fictitious retailer. The order book is generated; every figure on this sheet is a formula over it.',
		{ italic: true, color: P.muted, size: 9 },
	);
	book.wb.activeSheet = sheet.index;
}
