/** The balance sheet, with the check that turns red when it stops balancing. */
import type { Book } from '../lib/book';
import { K } from './costs';
import { R, revenueCell } from './revenue';
import { B, CASH, COSTS, F, I, INCOME, all, sum } from './rows';
import { FMT, P, SHEETS, YEARS, frame, line, section, yearHeader } from './style';

/** The opening balance sheet (FY2025 actuals); retained earnings is what makes it balance. */
const OPENING = {
	cash: 1840000,
	receivables: 119700,
	inventory: 193300,
	fixed: 9640000,
	payables: 665900,
	debt: 3200000,
	capital: 2000000,
};
const OPENING_RETAINED =
	OPENING.cash +
	OPENING.receivables +
	OPENING.inventory +
	OPENING.fixed -
	OPENING.payables -
	OPENING.debt -
	OPENING.capital;

export function buildBalance(book: Book): void {
	const sheet = book.sheet(SHEETS.balance);
	frame(
		sheet,
		'Balance sheet',
		'Dollars, shown in thousands. The base year column holds the opening position as typed actuals.',
		P.crust,
	);
	yearHeader(sheet, 'Line item', true);
	sheet.widths({ A: 2, B: 40, C: 10, D: 13, 'E:I': 12.5 }).freeze(5, 3);

	section(sheet, 8, 'Assets');
	line(sheet, B.cash, {
		label: 'Cash',
		unit: '$000',
		link: true,
		base: OPENING.cash,
		cell: (col) => `=${CASH}!${col}${F.closing}`,
	});
	line(sheet, B.receivables, {
		label: 'Trade receivables',
		unit: '$000',
		base: OPENING.receivables,
		cell: (col) => `=${revenueCell(col, R.wholesale)}*ReceivableDays/365`,
	});
	line(sheet, B.inventory, {
		label: 'Inventory',
		unit: '$000',
		base: OPENING.inventory,
		cell: (col) => `=${COSTS}!${col}${K.cogs}*InventoryDays/365`,
	});
	line(sheet, B.current, {
		label: 'Current assets',
		unit: '$000',
		kind: 'subtotal',
		base: sum(B.cash, B.inventory)('D'),
		cell: sum(B.cash, B.inventory),
	});
	sheet.groupRows(B.cash, B.inventory);
	line(sheet, B.fixed, {
		label: 'Property, plant and equipment',
		unit: '$000',
		kind: 'label',
		base: OPENING.fixed,
		cell: (col, prev) =>
			`=${prev}${B.fixed}+${COSTS}!${col}${K.capex}-${COSTS}!${col}${K.depreciation}`,
	});
	line(sheet, B.assets, {
		label: 'Total assets',
		unit: '$000',
		kind: 'total',
		...all((col) => `=${col}${B.current}+${col}${B.fixed}`),
	});

	section(sheet, 16, 'Liabilities');
	line(sheet, B.payables, {
		label: 'Trade payables',
		unit: '$000',
		base: OPENING.payables,
		cell: (col) => `=${COSTS}!${col}${K.cogs}*PayableDays/365`,
	});
	line(sheet, B.debt, {
		label: 'Bank debt',
		unit: '$000',
		base: OPENING.debt,
		cell: (col, prev) =>
			`=${prev}${B.debt}+${CASH}!${col}${F.borrowing}+${CASH}!${col}${F.repayments}`,
	});
	line(sheet, B.liabilities, {
		label: 'Total liabilities',
		unit: '$000',
		kind: 'subtotal',
		base: sum(B.payables, B.debt)('D'),
		cell: sum(B.payables, B.debt),
	});
	sheet.groupRows(B.payables, B.debt);

	section(sheet, 21, 'Equity');
	line(sheet, B.capital, {
		label: 'Share capital',
		unit: '$000',
		base: OPENING.capital,
		cell: (_col, prev) => `=${prev}${B.capital}`,
	});
	line(sheet, B.retained, {
		label: 'Retained earnings',
		unit: '$000',
		base: OPENING_RETAINED,
		cell: (col, prev) =>
			`=${prev}${B.retained}+${INCOME}!${col}${I.net}+${CASH}!${col}${F.dividends}`,
	});
	line(sheet, B.equity, {
		label: 'Total equity',
		unit: '$000',
		kind: 'subtotal',
		base: sum(B.capital, B.retained)('D'),
		cell: sum(B.capital, B.retained),
	});
	sheet.groupRows(B.capital, B.retained);
	line(sheet, B.funding, {
		label: 'Total liabilities and equity',
		unit: '$000',
		kind: 'total',
		...all((col) => `=${col}${B.liabilities}+${col}${B.equity}`),
	});

	sheet.set(`B${B.check}`, 'Balance check: total assets less total funding', {
		bold: true,
		v: 'center',
	});
	for (const col of ['D', ...YEARS])
		sheet.formula(`${col}${B.check}`, `ROUND(${col}${B.assets}-${col}${B.funding},2)`, {
			named: 'Check Cell',
			fmt: FMT.check,
		});
	// The formula rule turns the cell red as soon as a column stops balancing.
	sheet.conditional(`D${B.check}:I${B.check}`, {
		type: 'expression',
		formula: `D${B.check}<>0`,
		style: {
			font: { bold: true, color: { rgb: 'FFFFFFFF' } },
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: `FF${P.berry}` } },
		},
	});
	sheet.heights({ [B.check]: 19, [B.assets]: 19, [B.funding]: 19 });

	section(sheet, 29, 'Leverage');
	line(sheet, B.netDebt, {
		label: 'Net debt (bank debt less cash)',
		unit: '$000',
		kind: 'label',
		...all((col) => `=${col}${B.debt}-${col}${B.cash}`),
	});
	line(sheet, B.leverage, {
		label: 'Net debt to EBITDA',
		unit: 'multiple',
		kind: 'ratio',
		fmt: FMT.multiple,
		...all((col) => `=${col}${B.netDebt}/${INCOME}!${col}${I.ebitda}`),
	});
}
