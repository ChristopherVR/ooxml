/**
 * The Revenue build: the store estate, the retail drivers and each revenue stream by half year.
 * Half-year columns are outlined under the full-year column that sums them.
 */
import type { Book, Look } from '../lib/book';
import { L, type Sheet } from '../lib/sheet';
import { A } from './assumptions';
import { FMT, FOOTER, LOOK, P, SHEETS, YEARS, frame, section } from './style';

const S = SHEETS.revenue;
const ASSUME = SHEETS.assumptions;

/** Rows other sheets read. */
export const R = {
	opening: 9,
	newStores: 10,
	closing: 11,
	equivalents: 12,
	transactions: 15,
	ticket: 16,
	perStore: 17,
	established: 20,
	newRevenue: 21,
	retail: 22,
	wholesale: 23,
	online: 24,
	total: 25,
	growth: 26,
} as const;

interface YearColumns {
	h1: string;
	h2: string;
	fy: string;
	/** The full-year column of the year before (the base year for the first). */
	prev: string;
	/** The column of this year on the Assumptions sheet. */
	input: string;
}

const COLUMNS: YearColumns[] = YEARS.map((input, i) => ({
	h1: L(4 + 3 * i),
	h2: L(5 + 3 * i),
	fy: L(6 + 3 * i),
	prev: i === 0 ? 'D' : L(3 + 3 * i),
	input,
}));

/** The full-year column of forecast year `index` (0 to 4) on this sheet. */
export const revenueColumn = (index: number): string => COLUMNS[index]!.fy;
/** A full-year Revenue cell for the model column `col` (`D` is the base year). */
export const revenueCell = (col: string, row: number): string =>
	`${S}!${col === 'D' ? 'D' : revenueColumn(YEARS.indexOf(col as (typeof YEARS)[number]))}${row}`;

const LAST = COLUMNS[4]!.fy;

function header(sheet: Sheet): void {
	const head: Look = { bold: true, color: P.white, fill: P.ink, v: 'center' };
	sheet.paint(`B5:${LAST}6`, head);
	sheet.set('B5', 'Line item', { ...head, indent: 1 }).set('C5', 'Unit', head);
	sheet.formula('D5', `${ASSUME}!D5`, { ...head, h: 'right', fmt: FMT.actualYear });
	sheet.set('D6', 'Full year', { ...head, h: 'right', size: 9 });
	for (const year of COLUMNS) {
		sheet.formula(`${year.h1}5`, `${ASSUME}!${year.input}5`, {
			...head,
			h: 'center',
			fmt: FMT.forecastYear,
			left: P.muted,
		});
		sheet.merge(`${year.h1}5:${year.fy}5`);
		sheet.set(`${year.h1}6`, 'H1', { ...head, h: 'right', size: 9, left: P.muted });
		sheet.set(`${year.h2}6`, 'H2', { ...head, h: 'right', size: 9 });
		sheet.set(`${year.fy}6`, 'Full year', { ...head, h: 'right', size: 9 });
	}
}

type Kind = 'detail' | 'subtotal' | 'total' | 'ratio';
interface Row {
	label: string;
	unit: string;
	fmt: string;
	kind?: Kind;
	link?: boolean;
	base?: string;
	/** The full-year formula. */
	year: (y: YearColumns) => string;
	/** The half-year formula, when the line is built by half (`index` is the seasonality factor). */
	half?: (y: YearColumns, col: string, index: string) => string;
}

function write(sheet: Sheet, row: number, spec: Row): void {
	const base: Look =
		spec.kind === 'total'
			? { named: 'Total' }
			: spec.kind === 'subtotal'
				? LOOK.subtotal
				: spec.kind === 'ratio'
					? LOOK.ratio
					: { v: 'center' };
	if (spec.kind === 'total' || spec.kind === 'subtotal') sheet.paint(`B${row}:${LAST}${row}`, base);
	sheet.set(`B${row}`, spec.label, {
		...base,
		...(spec.kind === 'total' || spec.kind === 'subtotal' ? {} : { indent: 1 }),
	});
	sheet.set(`C${row}`, spec.unit, { ...base, ...LOOK.unit });
	const value: Look = { ...base, fmt: spec.fmt, ...(spec.link ? { color: P.linkInk } : {}) };
	// Full-year columns carry a tint so the eye finds the summary of each outline group.
	const summary: Look = { ...value, fill: P.cream };
	if (spec.base) sheet.formula(`D${row}`, spec.base, summary);
	else sheet.paint(`D${row}`, summary);
	for (const year of COLUMNS) {
		const edge: Look = { left: P.rule };
		if (spec.half) {
			sheet.formula(`${year.h1}${row}`, spec.half(year, year.h1, 'H1Index'), { ...value, ...edge });
			sheet.formula(`${year.h2}${row}`, spec.half(year, year.h2, '(2-H1Index)'), value);
		} else sheet.paint(`${year.h1}${row}`, { ...value, ...edge });
		sheet.formula(`${year.fy}${row}`, spec.year(year), summary);
	}
}

export function buildRevenue(book: Book): void {
	const sheet = book.sheet(S);
	frame(
		sheet,
		'Revenue build',
		'Store estate, retail drivers and revenue by stream. Dollars; half years fold under each full year (outline buttons above the columns).',
		P.sage,
	);
	sheet.page(FOOTER, { tall: 1 });
	header(sheet);
	sheet.heights({ 5: 20, 6: 17, 7: 8 });
	sheet.widths({ A: 2, B: 40, C: 10, D: 13 });
	for (const year of COLUMNS) {
		sheet.widths({ [`${year.h1}:${year.h2}`]: 12 }).groupColumns(`${year.h1}:${year.h2}`);
		sheet.widths({ [year.fy]: 13.5 });
	}
	sheet.freeze(6, 3);
	const sum = (y: YearColumns, row: number) => `SUM(${y.h1}${row}:${y.h2}${row})`;
	const halfYear = `TradingWeeks/2`;

	section(sheet, 8, 'Store estate', LAST);
	write(sheet, R.opening, {
		label: 'Stores open at the start of the year',
		unit: 'stores',
		fmt: FMT.count,
		year: (y) => `${y.prev}${R.closing}`,
	});
	write(sheet, R.newStores, {
		label: 'New store openings',
		unit: 'stores',
		fmt: FMT.count,
		link: true,
		year: (y) => `${ASSUME}!${y.input}${A.newStores}`,
	});
	write(sheet, R.closing, {
		label: 'Stores open at the end of the year',
		unit: 'stores',
		fmt: FMT.count,
		kind: 'subtotal',
		base: 'OpeningStores',
		year: (y) => `${y.fy}${R.opening}+${y.fy}${R.newStores}`,
	});
	write(sheet, R.equivalents, {
		label: 'Store equivalents trading',
		unit: 'stores',
		fmt: '#,##0.0',
		kind: 'ratio',
		base: `D${R.closing}`,
		year: (y) => `${y.fy}${R.opening}+${y.fy}${R.newStores}*NewStoreProductivity`,
	});
	sheet.groupRows(R.opening, R.newStores);

	section(sheet, 14, 'Retail drivers', LAST);
	write(sheet, R.transactions, {
		label: 'Weekly transactions per established store',
		unit: 'tickets',
		fmt: FMT.count,
		base: 'BaseWeeklyTransactions',
		year: (y) => `${y.prev}${R.transactions}*(1+${ASSUME}!${y.input}${A.likeForLike})`,
	});
	write(sheet, R.ticket, {
		label: 'Average ticket',
		unit: '$',
		fmt: FMT.price,
		base: 'BaseTicket',
		year: (y) => `${y.prev}${R.ticket}*(1+${ASSUME}!${y.input}${A.ticketInflation})`,
	});
	write(sheet, R.perStore, {
		label: 'Annual revenue of an established store',
		unit: '$',
		fmt: FMT.money,
		kind: 'subtotal',
		base: `D${R.transactions}*D${R.ticket}*TradingWeeks`,
		year: (y) => `${y.fy}${R.transactions}*${y.fy}${R.ticket}*TradingWeeks`,
	});

	section(sheet, 19, 'Revenue', LAST);
	write(sheet, R.established, {
		label: 'Established stores',
		unit: '$',
		fmt: FMT.money,
		base: `D${R.equivalents}*D${R.perStore}`,
		half: (y, _col, index) =>
			`${y.fy}${R.opening}*${y.fy}${R.transactions}*${y.fy}${R.ticket}*${halfYear}*${index}`,
		year: (y) => sum(y, R.established),
	});
	write(sheet, R.newRevenue, {
		label: 'New stores in their opening year',
		unit: '$',
		fmt: FMT.money,
		half: (y, _col, index) =>
			`${y.fy}${R.newStores}*NewStoreProductivity*${y.fy}${R.transactions}*${y.fy}${R.ticket}*${halfYear}*${index}`,
		year: (y) => sum(y, R.newRevenue),
	});
	write(sheet, R.retail, {
		label: 'Retail stores',
		unit: '$',
		fmt: FMT.money,
		kind: 'subtotal',
		base: `SUM(D${R.established}:D${R.newRevenue})`,
		half: (_y, col) => `SUM(${col}${R.established}:${col}${R.newRevenue})`,
		year: (y) => `SUM(${y.fy}${R.established}:${y.fy}${R.newRevenue})`,
	});
	write(sheet, R.wholesale, {
		label: 'Wholesale',
		unit: '$',
		fmt: FMT.money,
		base: 'WholesaleBase',
		half: (y) => `${y.prev}${R.wholesale}*(1+${ASSUME}!${y.input}${A.wholesaleGrowth})/2`,
		year: (y) => sum(y, R.wholesale),
	});
	write(sheet, R.online, {
		label: 'Online orders',
		unit: '$',
		fmt: FMT.money,
		base: 'OnlineBase',
		half: (y, _col, index) =>
			`${y.prev}${R.online}*(1+${ASSUME}!${y.input}${A.onlineGrowth})/2*${index}`,
		year: (y) => sum(y, R.online),
	});
	write(sheet, R.total, {
		label: 'Total revenue',
		unit: '$',
		fmt: FMT.money,
		kind: 'total',
		base: `D${R.retail}+D${R.wholesale}+D${R.online}`,
		half: (_y, col) => `${col}${R.retail}+${col}${R.wholesale}+${col}${R.online}`,
		year: (y) => `${y.fy}${R.retail}+${y.fy}${R.wholesale}+${y.fy}${R.online}`,
	});
	write(sheet, R.growth, {
		label: 'Growth on the prior year',
		unit: '%',
		fmt: FMT.percent,
		kind: 'ratio',
		year: (y) => `${y.fy}${R.total}/${y.prev}${R.total}-1`,
	});
	// Two outline levels: the retail detail folds into Retail stores, the streams into the total.
	sheet.groupRows(R.retail, R.online).groupRows(R.established, R.newRevenue, 2);
	sheet.heights({ [R.total]: 19 });

	section(sheet, 28, 'Revenue mix', LAST);
	const share = (label: string, row: number, source: number) =>
		write(sheet, row, {
			label,
			unit: '% of total',
			fmt: FMT.percent,
			kind: 'ratio',
			base: `D${source}/D${R.total}`,
			year: (y) => `${y.fy}${source}/${y.fy}${R.total}`,
		});
	share('Retail stores', 29, R.retail);
	share('Wholesale', 30, R.wholesale);
	share('Online orders', 31, R.online);
}
