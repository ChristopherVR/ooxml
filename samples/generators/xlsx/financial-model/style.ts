/** The look of the Brindlemoor Bakehouse model: palette, number formats, cell styles, sheet frame. */
import { Book, type Look } from '../lib/book';
import type { Sheet } from '../lib/sheet';

/** A warm bakery palette: espresso ink, crust and honey accents on cream paper. */
export const P = {
	ink: '2B2118',
	muted: '7A6A5B',
	crust: '9C4A26',
	honey: 'D9A441',
	sage: '6B7F4E',
	teal: '3E6B73',
	berry: 'A23B3B',
	walnut: '8C5A3C',
	cream: 'FBF6EE',
	sand: 'F3E9DA',
	rule: 'DCCFBD',
	inputFill: 'FFF4D6',
	inputInk: '1F5F7A',
	linkInk: '3F6B3A',
	white: 'FFFFFF',
} as const;

export const FMT = {
	/** Whole currency units shown in thousands, negatives in parentheses, zero as a dash. */
	thousands: '#,##0,_);(#,##0,);"-"_)',
	money: '"$"#,##0_);("$"#,##0);"-"_)',
	price: '"$"#,##0.00',
	millions: '"$"#,##0.0,,"m"',
	percent: '0.0%_);(0.0%);"-"_)',
	count: '#,##0_);(#,##0);"-"_)',
	index: '0.00',
	days: '0" days"',
	years: '0" years"',
	multiple: '0.0"x"_);(0.0"x")',
	forecastYear: '"FY"0"E"',
	actualYear: '"FY"0"A"',
	check: '[=0]"Balanced";[Red]"Out by "#,##0.00',
} as const;

/** The left footer of every printed page. */
export const FOOTER = 'Brindlemoor Bakehouse five-year model';

export const SHEETS = {
	cover: 'Cover',
	assumptions: 'Assumptions',
	revenue: 'Revenue',
	costs: 'Costs',
	income: 'Income statement',
	balance: 'Balance sheet',
	cash: 'Cash flow',
	dashboard: 'Dashboard',
} as const;

/** The five forecast columns, shared by every model sheet except the half-year revenue build. */
export const YEARS = ['E', 'F', 'G', 'H', 'I'] as const;
/** The column before a forecast column (the base year for the first). */
export const before = (col: string): string => String.fromCharCode(col.charCodeAt(0) - 1);

export const LOOK = {
	label: { v: 'center' } as Look,
	detail: { indent: 1, v: 'center' } as Look,
	unit: { color: P.muted, size: 9, v: 'center' } as Look,
	subtotal: { bold: true, top: P.rule, v: 'center' } as Look,
	ratio: { italic: true, color: P.muted, v: 'center' } as Look,
	note: { named: 'Explanatory Text' } as Look,
};

export function createBook(): Book {
	const book = new Book(
		Object.values(SHEETS),
		{
			light1: P.white,
			dark1: P.ink,
			light2: P.sand,
			dark2: '5A3E2B',
			accents: [P.crust, P.honey, P.sage, P.teal, P.walnut, P.berry],
			link: P.inputInk,
			followedLink: P.walnut,
		},
		P.ink,
	);
	// The gallery of Home > Cell Styles: Excel's built-in names, restyled for this model.
	book.cellStyle('Title', { font: 'display', size: 20, bold: true, color: P.crust }, 15);
	book.cellStyle('Heading 1', { font: 'display', size: 16, bold: true, color: P.crust }, 16);
	book.cellStyle(
		'Heading 2',
		{ font: 'display', size: 10.5, bold: true, color: P.crust, bottom: P.crust, v: 'bottom' },
		17,
	);
	book.cellStyle(
		'Input',
		{ color: P.inputInk, fill: P.inputFill, bottom: P.honey, top: P.honey, unlocked: true },
		20,
	);
	book.cellStyle('Calculation', { color: P.ink }, 22);
	book.cellStyle('Linked Cell', { color: P.linkInk }, 24);
	book.cellStyle('Total', { bold: true, top: P.ink, bottom: ['double', P.ink] }, 25);
	book.cellStyle('Check Cell', { bold: true, color: P.white, fill: P.sage, h: 'center' }, 23);
	book.cellStyle('Explanatory Text', { italic: true, color: P.muted, size: 9 }, 53);
	return book;
}

/** The title block and the header row every model sheet shares; returns nothing to keep. */
export function frame(sheet: Sheet, title: string, subtitle: string, tab: string): void {
	// Statements print upright on one page; wider sheets override this.
	sheet.present(tab).page(FOOTER, { portrait: true, tall: 1 });
	sheet.set('B2', title, { named: 'Heading 1' });
	sheet.set('B3', subtitle, LOOK.note);
	sheet.heights({ 1: 9, 2: 27, 3: 16, 4: 9, 5: 21, 6: 6, 7: 6 });
}

/** The dark header row with the base year and the five forecast years. */
export function yearHeader(sheet: Sheet, first: string, linked: boolean): void {
	const head: Look = { bold: true, color: P.white, fill: P.ink, v: 'center' };
	sheet.paint('B5:I5', head);
	sheet.set('B5', first, { ...head, indent: 1 });
	sheet.set('C5', 'Unit', head);
	const year: Look = { ...head, h: 'right' };
	if (linked) {
		sheet.formula('D5', `${SHEETS.assumptions}!D5`, { ...year, fmt: FMT.actualYear });
		for (const col of YEARS)
			sheet.formula(`${col}5`, `${SHEETS.assumptions}!${col}5`, {
				...year,
				fmt: FMT.forecastYear,
			});
	}
}

export function section(sheet: Sheet, row: number, text: string, to = 'I'): void {
	sheet.paint(`B${row}:${to}${row}`, { named: 'Heading 2' });
	sheet.set(`B${row}`, text);
	sheet.heights({ [row]: 24 });
}

/** One line of a statement: its label, unit, base-year cell and the five forecast cells. */
export interface Line {
	label: string;
	unit?: string;
	/** The base-year cell: a formula (`=...`), a typed actual, or nothing. */
	base?: string | number;
	/** The forecast cell of column `col`, with `prev` the column before it. */
	cell?: (col: string, prev: string, index: number) => string | number;
	look?: Look;
	/** Number format of the value cells. */
	fmt?: string;
	/** `link` marks cells that only pull a value from another sheet. */
	kind?: 'detail' | 'subtotal' | 'total' | 'ratio' | 'label';
	link?: boolean;
	/** Typed values are inputs. */
	input?: boolean;
}

function valueLook(line: Line, value: string | number): Look {
	const fmt = line.fmt ?? FMT.thousands;
	const typed = typeof value === 'number';
	if (line.kind === 'total') return { named: 'Total', fmt };
	if (line.kind === 'subtotal') return { ...LOOK.subtotal, fmt };
	if (line.kind === 'ratio') return { ...LOOK.ratio, fmt };
	if (typed || line.input) return { named: 'Input', fmt };
	if (line.link) return { named: 'Linked Cell', fmt };
	return { fmt, ...line.look };
}

/** Writes a line at `row`; a formula starts with `=`. */
export function line(sheet: Sheet, row: number, spec: Line): void {
	const labelLook =
		spec.kind === 'total'
			? { named: 'Total' }
			: spec.kind === 'subtotal'
				? LOOK.subtotal
				: spec.kind === 'ratio'
					? { ...LOOK.ratio, indent: 1 }
					: spec.kind === 'label'
						? LOOK.label
						: LOOK.detail;
	if (spec.kind === 'total' || spec.kind === 'subtotal') sheet.paint(`B${row}:I${row}`, labelLook);
	sheet.set(`B${row}`, spec.label, labelLook);
	if (spec.unit) sheet.set(`C${row}`, spec.unit, { ...labelLook, ...LOOK.unit, indent: 0 });
	const put = (ref: string, value: string | number) => {
		const look = valueLook(spec, value);
		if (typeof value === 'string' && value.startsWith('='))
			sheet.formula(ref, value.slice(1), look);
		else sheet.set(ref, value, look);
	};
	if (spec.base !== undefined) put(`D${row}`, spec.base);
	if (spec.cell) YEARS.forEach((col, i) => put(`${col}${row}`, spec.cell!(col, before(col), i)));
}
