/** The Cover sheet: contents with links, the colour key of the model and its integrity checks. */
import type { Book, Look } from '../lib/book';
import { PEOPLE } from '../../lib/fiction';
import { FOOTER, LOOK, P, SHEETS } from './style';

const CONTENTS: [string, string][] = [
	[SHEETS.assumptions, 'Scenario selector and every input, each with a validation rule'],
	[SHEETS.revenue, 'Store estate, retail drivers and revenue by stream, by half year'],
	[SHEETS.costs, 'Cost of goods sold, store costs, overheads, capital expenditure'],
	[SHEETS.income, 'Revenue to net income, with margins'],
	[SHEETS.balance, 'Assets, liabilities and equity, with the balance check'],
	[SHEETS.cash, 'Operating, investing and financing cash flows, reconciled to cash'],
	[SHEETS.dashboard, 'Headline figures, the five-year summary and charts'],
];

export function buildCover(book: Book): void {
	const sheet = book.sheet(SHEETS.cover);
	sheet.present(P.ink).page(FOOTER, { portrait: true, tall: 1 });
	sheet.widths({ A: 3, B: 5, C: 25, D: 64, E: 3, F: 26 });
	sheet.heights({ 1: 14, 2: 16, 3: 36, 4: 20, 5: 8 });
	sheet.set('B2', 'BRINDLEMOOR BAKEHOUSE', { bold: true, size: 9, color: P.muted });
	sheet.set('B3', 'Five-year financial model', { named: 'Title' });
	sheet.set('B4', 'Plan for FY2026 to FY2030, prepared for the board meeting of 12 February 2026', {
		size: 11,
		color: P.muted,
	});
	sheet.paint('B5:F5', { bottom: ['medium', P.crust] });

	sheet.paint('B7:D7', { named: 'Heading 2' }).set('B7', 'Contents');
	const link: Look = { color: P.inputInk, underline: true, v: 'center' };
	CONTENTS.forEach(([name, what], i) => {
		const row = 8 + i;
		sheet.paint(`B${row}:D${row}`, { bottom: P.sand, v: 'center' });
		sheet.set(`B${row}`, i + 1, { color: P.muted, h: 'center', bottom: P.sand, v: 'center' });
		sheet.set(`C${row}`, name, { ...link, bottom: P.sand });
		sheet.set(`D${row}`, what, { bottom: P.sand, v: 'center' });
		sheet.linkTo(`C${row}`, `'${name}'!A1`, `Go to ${name}`);
		sheet.heights({ [row]: 18 });
	});

	sheet.paint('B17:D17', { named: 'Heading 2' }).set('B17', 'Colour key');
	const key: [Look, string, string][] = [
		[
			{ named: 'Input', h: 'center' },
			'Input',
			'Typed by hand and validated. The only cells to change.',
		],
		[
			{ named: 'Calculation', h: 'center' },
			'Calculation',
			'A formula over cells of the same sheet.',
		],
		[{ named: 'Linked Cell', h: 'center' }, 'Link', 'A value pulled from another sheet.'],
		[{ named: 'Total', h: 'center' }, 'Total', 'A total or subtotal of the lines above it.'],
		[
			{ named: 'Check Cell' },
			'Check',
			'An integrity check. It turns red when the model stops tying.',
		],
	];
	key.forEach(([look, label, what], i) => {
		const row = 18 + i;
		sheet.set(`C${row}`, label, { ...look, v: 'center' });
		sheet.set(`D${row}`, what, { v: 'center', indent: 1 });
		sheet.heights({ [row]: 18 });
	});

	sheet.paint('B24:D24', { named: 'Heading 2' }).set('B24', 'Model integrity');
	const checks: [string, string][] = [
		[
			'Balance sheet',
			`IF(COUNTIF('${SHEETS.balance}'!Check,"<>0")=0,"Balances in every year","Out of balance: see the check row")`,
		],
		[
			'Cash flow',
			`IF(COUNTIF('${SHEETS.cash}'!Check,"<>0")=0,"Reconciles to balance sheet cash","Does not reconcile: see the check row")`,
		],
		[
			'Cash floor',
			'IF(MIN(Closing_Cash)>=MinimumCash,"Cash stays above the minimum balance","Cash falls below the minimum balance")',
		],
	];
	checks.forEach(([label, formula], i) => {
		const row = 25 + i;
		sheet.set(`C${row}`, label, { v: 'center' });
		sheet.formula(`D${row}`, formula, { bold: true, color: P.sage, v: 'center', indent: 1 });
		sheet.heights({ [row]: 18 });
	});
	// A failed check starts with "Out", "Does not" or "Cash falls": paint those red.
	sheet.conditional(
		'D25:D27',
		...['Out of', 'Does not', 'Cash falls'].map((text) => ({
			type: 'beginsWith' as const,
			text,
			style: { font: { bold: true, color: { rgb: `FF${P.berry}` } } },
		})),
	);

	sheet.paint('B29:D29', { named: 'Heading 2' }).set('B29', 'About this workbook');
	const about: [string, string][] = [
		['Prepared by', `${PEOPLE[0]}, Finance`],
		['Reviewed by', `${PEOPLE[1]}, Chief Financial Officer`],
		['Version', '1.4, 15 January 2026'],
		['Units', 'Dollars. Statements are shown in thousands; inputs are in whole dollars.'],
	];
	about.forEach(([label, text], i) => {
		sheet.set(`C${30 + i}`, label, { color: P.muted, v: 'center' });
		sheet.set(`D${30 + i}`, text, { v: 'center', indent: 1 });
	});
	sheet.set('C34', 'Model guide', { color: P.muted, v: 'center' });
	sheet.set('D34', 'finance.brindlemoor.example.invalid/planning/model-guide', {
		...link,
		indent: 1,
	});
	sheet.linkOut(
		'D34',
		'https://finance.brindlemoor.example.invalid/planning/model-guide',
		'The planning team guide to this model',
	);
	sheet.set(
		'B36',
		'Brindlemoor Bakehouse is a fictitious company. Every figure in this workbook is invented for demonstration.',
		LOOK.note,
	);
	book.wb.activeSheet = sheet.index;
}
