/** The cash flow, derived from net income and the movements of the balance sheet. */
import type { Book } from '../lib/book';
import { A } from './assumptions';
import { K } from './costs';
import { ASSUME, B, BALANCE, COSTS, F, I, INCOME, sum } from './rows';
import { P, SHEETS, YEARS, before, frame, line, section, yearHeader } from './style';

export function buildCashFlow(book: Book): void {
	const sheet = book.sheet(SHEETS.cash);
	frame(
		sheet,
		'Cash flow',
		'Dollars, shown in thousands. Indirect method: net income, working capital movements, investment, financing.',
		P.crust,
	);
	yearHeader(sheet, 'Line item', true);
	sheet.widths({ A: 2, B: 40, C: 10, D: 13, 'E:I': 12.5 }).freeze(5, 3);
	const movement = (row: number, asset: boolean) => (col: string, prev: string) =>
		asset
			? `=${BALANCE}!${prev}${row}-${BALANCE}!${col}${row}`
			: `=${BALANCE}!${col}${row}-${BALANCE}!${prev}${row}`;

	section(sheet, 8, 'Operating activities');
	line(sheet, F.net, {
		label: 'Net income',
		unit: '$000',
		link: true,
		cell: (col) => `=${INCOME}!${col}${I.net}`,
	});
	line(sheet, F.depreciation, {
		label: 'Depreciation added back',
		unit: '$000',
		link: true,
		cell: (col) => `=${COSTS}!${col}${K.depreciation}`,
	});
	line(sheet, F.receivables, {
		label: 'Decrease (increase) in receivables',
		unit: '$000',
		cell: movement(B.receivables, true),
	});
	line(sheet, F.inventory, {
		label: 'Decrease (increase) in inventory',
		unit: '$000',
		cell: movement(B.inventory, true),
	});
	line(sheet, F.payables, {
		label: 'Increase (decrease) in payables',
		unit: '$000',
		cell: movement(B.payables, false),
	});
	line(sheet, F.operations, {
		label: 'Cash from operating activities',
		unit: '$000',
		kind: 'subtotal',
		cell: sum(F.net, F.payables),
	});
	sheet.groupRows(F.net, F.payables);

	section(sheet, 16, 'Investing activities');
	line(sheet, F.fitOut, {
		label: 'New store fit-out',
		unit: '$000',
		link: true,
		cell: (col) => `=-${COSTS}!${col}${K.fitOut}`,
	});
	line(sheet, F.maintenance, {
		label: 'Maintenance of the estate',
		unit: '$000',
		link: true,
		cell: (col) => `=-${COSTS}!${col}${K.maintenance}`,
	});
	line(sheet, F.extension, {
		label: 'Central bakery extension',
		unit: '$000',
		link: true,
		cell: (col) => `=-${COSTS}!${col}${K.extension}`,
	});
	line(sheet, F.investing, {
		label: 'Cash from investing activities',
		unit: '$000',
		kind: 'subtotal',
		cell: sum(F.fitOut, F.extension),
	});
	sheet.groupRows(F.fitOut, F.extension);

	section(sheet, 22, 'Financing activities');
	line(sheet, F.borrowing, {
		label: 'New borrowing drawn',
		unit: '$000',
		link: true,
		cell: (col) => `=${ASSUME}!${col}${A.borrowing}`,
	});
	line(sheet, F.repayments, {
		label: 'Debt repaid',
		unit: '$000',
		cell: (col, prev) =>
			`=-MIN(${ASSUME}!${col}${A.repayments},${BALANCE}!${prev}${B.debt}+${col}${F.borrowing})`,
	});
	line(sheet, F.dividends, {
		label: 'Dividends paid',
		unit: '$000',
		cell: (col) => `=-MAX(0,${INCOME}!${col}${I.net})*PayoutRatio`,
	});
	line(sheet, F.financing, {
		label: 'Cash from financing activities',
		unit: '$000',
		kind: 'subtotal',
		cell: sum(F.borrowing, F.dividends),
	});
	sheet.groupRows(F.borrowing, F.dividends);

	line(sheet, F.change, {
		label: 'Net change in cash',
		unit: '$000',
		kind: 'subtotal',
		cell: (col) => `=${col}${F.operations}+${col}${F.investing}+${col}${F.financing}`,
	});
	line(sheet, F.opening, {
		label: 'Cash at the start of the year',
		unit: '$000',
		kind: 'label',
		cell: (_col, prev) => `=${prev}${F.closing}`,
	});
	line(sheet, F.closing, {
		label: 'Cash at the end of the year',
		unit: '$000',
		kind: 'total',
		base: `=${BALANCE}!D${B.cash}`,
		cell: (col) => `=${col}${F.change}+${col}${F.opening}`,
	});
	sheet.heights({ [F.closing]: 19 });

	section(sheet, 32, 'Checks');
	sheet.set(`B${F.check}`, 'Movement in balance sheet cash less net cash flow', {
		bold: true,
		v: 'center',
	});
	for (const col of YEARS)
		sheet.formula(
			`${col}${F.check}`,
			`ROUND(${BALANCE}!${col}${B.cash}-${BALANCE}!${before(col)}${B.cash}-${col}${F.change},2)`,
			{ named: 'Check Cell', fmt: '[=0]"Reconciled";[Red]"Out by "#,##0.00' },
		);
	sheet.conditional(`E${F.check}:I${F.check}`, {
		type: 'expression',
		formula: `E${F.check}<>0`,
		style: {
			font: { bold: true, color: { rgb: 'FFFFFFFF' } },
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: `FF${P.berry}` } },
		},
	});
	line(sheet, F.headroom, {
		label: 'Headroom over the minimum cash balance',
		unit: '$000',
		kind: 'label',
		cell: (col) => `=${col}${F.closing}-MinimumCash`,
	});
	// Traffic lights: red below the floor, amber within a million of it, green above.
	sheet.conditional(`E${F.headroom}:I${F.headroom}`, {
		type: 'iconSet',
		iconSet: '3TrafficLights1',
		thresholds: [
			{ type: 'percent', value: '0' },
			{ type: 'num', value: '0' },
			{ type: 'num', value: '1000000' },
		],
	});
}
