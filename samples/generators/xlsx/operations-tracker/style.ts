/** The look of the Vale Orrin Logistics tracker: navy and steel with amber and signal accents. */
import { Book, type Look } from '../lib/book';
import type { Sheet } from '../lib/sheet';

export const P = {
	ink: '1B2330',
	muted: '667085',
	navy: '233A5E',
	blue: '2F5D9B',
	amber: 'E79A2B',
	teal: '3E9C8F',
	steel: '8A94A6',
	signal: 'C8452F',
	paper: 'F3F5F9',
	mist: 'E4E8F0',
	rule: 'CDD3DE',
	inputFill: 'FFF6DC',
	white: 'FFFFFF',
} as const;

export const FMT = {
	date: 'd mmm yyyy',
	shortDate: 'd mmm',
	month: 'mmm yy',
	count: '#,##0',
	hours: '#,##0.0',
	money: '"$"#,##0',
	percent: '0.0%',
	days: '0.0',
	/** Days overdue: nothing for work that is on time. */
	overdue: '0" d";;',
} as const;

export const SHEETS = {
	overview: 'Overview',
	log: 'Work orders',
	report: 'Depot report',
	depots: 'Depots',
} as const;

/** The date the tracker reports as of (15 January 2026), an input on the Overview sheet. */
export const AS_OF = 46037;

/** The left footer of every printed page. */
export const FOOTER = 'Vale Orrin Logistics operations tracker';

export const HEAD: Look = { bold: true, color: P.white, fill: P.navy, v: 'center' };
export const INPUT: Look = {
	fill: P.inputFill,
	color: P.blue,
	bold: true,
	top: P.amber,
	bottom: P.amber,
	left: P.amber,
	right: P.amber,
	unlocked: true,
	v: 'center',
};
export const TOTAL: Look = { bold: true, top: P.ink, bottom: ['double', P.ink], v: 'center' };
export const LINE: Look = { bottom: P.mist, v: 'center' };

export function createBook(): Book {
	return new Book(
		Object.values(SHEETS),
		{
			light1: P.white,
			dark1: P.ink,
			light2: P.mist,
			dark2: P.navy,
			accents: [P.blue, P.amber, P.teal, P.steel, P.signal, '6C5BA7'],
			link: P.blue,
			followedLink: '6C5BA7',
		},
		P.ink,
	);
}

export function title(sheet: Sheet, text: string, subtitle: string, tab: string): void {
	sheet.present(tab).page(FOOTER);
	sheet.set('B2', text, { font: 'display', size: 16, bold: true, color: P.navy });
	sheet.set('B3', subtitle, { color: P.muted, size: 9, italic: true });
	sheet.heights({ 1: 9, 2: 27, 3: 16 });
}

export function section(sheet: Sheet, row: number, from: string, to: string, text: string): void {
	sheet.paint(`${from}${row}:${to}${row}`, {
		font: 'display',
		size: 10.5,
		bold: true,
		color: P.navy,
		bottom: P.navy,
		v: 'center',
	});
	sheet.set(`${from}${row}`, text);
	sheet.heights({ [row]: 20 });
}

/** A header row: the names from `from` rightwards, text columns left and the rest right. */
export function header(sheet: Sheet, row: number, cells: [string, string, Look['h']?][]): void {
	for (const [col, text, h] of cells)
		sheet.set(`${col}${row}`, text, {
			...HEAD,
			h: h ?? 'right',
			...(h === 'left' ? { indent: 1 } : {}),
		});
	sheet.heights({ [row]: 20 });
}
