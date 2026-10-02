/**
 * File > New templates, generated from code through the core's edit session (no bundled files):
 * a monthly budget, a to-do list and an invoice, each with working formulas.
 */
import {
	createEditSession,
	createWorkbook,
	parseRange,
	type CellRange,
	type EditSession,
	type StylePatch,
	type Workbook,
} from '@christophervr/xlsx-core';

export type TemplateId = 'budget' | 'todo' | 'invoice';
type T = (key: string, vars?: Record<string, string | number>) => string;

export const TEMPLATES: ReadonlyArray<{ id: TemplateId; label: string; description: string }> = [
	{ id: 'budget', label: 'Monthly budget', description: 'Income and expenses with totals' },
	{ id: 'todo', label: 'To-do list', description: 'Tasks with due dates and status' },
	{ id: 'invoice', label: 'Invoice', description: 'Line items, tax and total due' },
];

const range = (ref: string): CellRange => parseRange(ref)!;
const TITLE: StylePatch = { font: { size: 18, bold: true, color: { theme: 3 } } };
const HEADER: StylePatch = {
	font: { bold: true, color: { theme: 0 } },
	fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: '217346' } },
};
const TOTAL: StylePatch = {
	font: { bold: true },
	border: { top: { style: 'thin' }, bottom: { style: 'double' } },
};
const MONEY = '#,##0.00';

function rows(session: EditSession, at: number, lines: (string | number)[][]): void {
	lines.forEach((line, r) =>
		line.forEach((value, c) => session.setCellInput(0, at + r, c, String(value))),
	);
}

function budget(session: EditSession, t: T): void {
	session.renameSheet(0, t('Budget'));
	session.setCellInput(0, 0, 0, t('Monthly budget'));
	session.applyStyle(0, [range('A1')], TITLE);
	rows(session, 2, [
		[t('Category'), t('Planned'), t('Actual'), t('Difference')],
		[t('Salary'), 4200, 4200],
		[t('Rent'), -1400, -1400],
		[t('Groceries'), -450, -512.3],
		[t('Utilities'), -180, -164.9],
		[t('Transport'), -120, -98],
		[t('Savings'), -600, -600],
	]);
	for (let r = 3; r <= 8; r++) session.setCellInput(0, r, 3, `=C${r + 1}-B${r + 1}`);
	session.setCellInput(0, 9, 0, t('Total'));
	for (const [c, col] of [
		[1, 'B'],
		[2, 'C'],
		[3, 'D'],
	] as const)
		session.setCellInput(0, 9, c, `=SUM(${col}4:${col}9)`);
	session.applyStyle(0, [range('A3:D3')], HEADER);
	session.applyStyle(0, [range('B4:D10')], { numFmt: MONEY });
	session.applyStyle(0, [range('A10:D10')], TOTAL);
	session.setColumnWidth(0, [0], 18);
	session.setColumnWidth(0, [1, 2, 3], 13);
}

function todo(session: EditSession, t: T): void {
	session.renameSheet(0, t('Tasks'));
	session.setCellInput(0, 0, 0, t('To-do list'));
	session.applyStyle(0, [range('A1')], TITLE);
	rows(session, 2, [
		[t('Task'), t('Due date'), t('Priority'), t('Status')],
		[t('Plan the week'), '2026-01-05', t('High'), t('Done')],
		[t('Send the report'), '2026-01-07', t('High'), t('In progress')],
		[t('Book travel'), '2026-01-12', t('Medium'), t('Not started')],
		[t('Review the budget'), '2026-01-15', t('Low'), t('Not started')],
	]);
	session.setCellInput(0, 8, 0, t('Open tasks'));
	session.setCellInput(0, 8, 1, `=COUNTIF(D4:D7,"<>${t('Done').replace(/"/g, '""')}")`);
	session.applyStyle(0, [range('A3:D3')], HEADER);
	session.applyStyle(0, [range('B4:B7')], { numFmt: 'yyyy-mm-dd' });
	session.applyStyle(0, [range('A9:B9')], { font: { bold: true } });
	session.setColumnWidth(0, [0], 24);
	session.setColumnWidth(0, [1, 2, 3], 14);
}

function invoice(session: EditSession, t: T): void {
	session.renameSheet(0, t('Invoice'));
	session.setCellInput(0, 0, 0, t('Invoice'));
	session.applyStyle(0, [range('A1')], TITLE);
	rows(session, 2, [
		[t('Invoice number'), 'INV-0001'],
		[t('Date'), '2026-01-31'],
		[t('Bill to'), t('Customer name')],
	]);
	rows(session, 6, [
		[t('Description'), t('Quantity'), t('Unit price'), t('Amount')],
		[t('Consulting (hours)'), 12, 85],
		[t('Design work'), 1, 950],
		[t('Hosting (months)'), 3, 25],
	]);
	for (let r = 7; r <= 9; r++) session.setCellInput(0, r, 3, `=B${r + 1}*C${r + 1}`);
	rows(session, 11, [
		['', '', t('Subtotal')],
		['', '', t('Tax rate')],
		['', '', t('Tax')],
		['', '', t('Total due')],
	]);
	session.setCellInput(0, 11, 3, '=SUM(D8:D10)');
	session.setCellInput(0, 12, 3, '15%');
	session.setCellInput(0, 13, 3, '=D12*D13');
	session.setCellInput(0, 14, 3, '=D12+D14');
	session.applyStyle(0, [range('B4')], { numFmt: 'yyyy-mm-dd', alignment: { horizontal: 'left' } });
	session.applyStyle(0, [range('A7:D7')], HEADER);
	session.applyStyle(0, [range('C8:D10'), range('D12'), range('D14:D15')], { numFmt: MONEY });
	session.applyStyle(0, [range('C15:D15')], TOTAL);
	session.setColumnWidth(0, [0], 26);
	session.setColumnWidth(0, [1, 2, 3], 13);
}

/** A new workbook filled from a template, with the labels in the display language. */
export function createTemplateWorkbook(id: TemplateId, t: T): Workbook {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	if (id === 'budget') budget(session, t);
	else if (id === 'todo') todo(session, t);
	else invoice(session, t);
	return session.workbook;
}
