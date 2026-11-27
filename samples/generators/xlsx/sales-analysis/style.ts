/** The look of the Kestrelline Outfitters workbook: an outdoor palette of pine, ember and lake. */
import { Book, type Look } from '../lib/book';
import type { Sheet } from '../lib/sheet';

export const P = {
	ink: '1D2A2F',
	muted: '66777D',
	pine: '2F6F62',
	ember: 'E08A3C',
	lake: '4C8DAE',
	lichen: 'C9B458',
	heather: '7A5C8E',
	rust: 'B8493E',
	paper: 'F4F7F3',
	mist: 'E3EBE6',
	rule: 'CFD9D3',
	white: 'FFFFFF',
} as const;

export const FMT = {
	money: '"$"#,##0',
	cents: '"$"#,##0.00',
	thousands: '"$"#,##0.0,"k"',
	percent: '0.0%',
	wholePercent: '0%',
	count: '#,##0',
	date: 'd mmm yyyy',
	month: 'mmm',
	monthYear: 'mmm yyyy',
	delta: '+0.0%;-0.0%;0.0%',
} as const;

export const SHEETS = {
	dashboard: 'Dashboard',
	explorer: 'Explorer',
	analysis: 'Analysis',
	orders: 'Orders',
	lookups: 'Lookups',
} as const;

/** The left footer of every printed page. */
export const FOOTER = 'Kestrelline Outfitters sales analysis 2025';

export const HEAD: Look = { bold: true, color: P.white, fill: P.pine, v: 'center' };

export function createBook(): Book {
	return new Book(
		Object.values(SHEETS),
		{
			light1: P.white,
			dark1: P.ink,
			light2: P.mist,
			dark2: '23404A',
			accents: [P.pine, P.ember, P.lake, P.lichen, P.heather, P.rust],
			link: P.lake,
			followedLink: P.heather,
		},
		P.ink,
	);
}

/** Title and subtitle of a sheet, with the tab colour and gridlines off. */
export function title(sheet: Sheet, text: string, subtitle: string, tab: string): void {
	sheet.present(tab).page(FOOTER);
	sheet.set('B2', text, { font: 'display', size: 16, bold: true, color: P.pine });
	sheet.set('B3', subtitle, { color: P.muted, size: 9, italic: true });
	sheet.heights({ 1: 9, 2: 27, 3: 16 });
}

/** A section heading with a rule under it, across `from`..`to` of one row. */
export function section(sheet: Sheet, row: number, from: string, to: string, text: string): void {
	sheet.paint(`${from}${row}:${to}${row}`, {
		font: 'display',
		size: 10.5,
		bold: true,
		color: P.pine,
		bottom: P.pine,
		v: 'center',
	});
	sheet.set(`${from}${row}`, text);
	sheet.heights({ [row]: 20 });
}
