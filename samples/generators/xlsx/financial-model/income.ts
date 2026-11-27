/** The income statement: revenue to net income, costs as negatives. */
import type { Book } from '../lib/book';
import { K } from './costs';
import { R, revenueCell } from './revenue';
import { B, BALANCE, COSTS, I, all, sum } from './rows';
import { FMT, P, SHEETS, frame, line, section, yearHeader } from './style';

export function buildIncome(book: Book): void {
	const sheet = book.sheet(SHEETS.income);
	frame(
		sheet,
		'Income statement',
		'Dollars, shown in thousands. Costs are negative and print in parentheses.',
		P.crust,
	);
	yearHeader(sheet, 'Line item', true);
	sheet.widths({ A: 2, B: 40, C: 10, D: 13, 'E:I': 12.5 }).freeze(5, 3);
	const cost = (row: number) => all((col) => `=-${COSTS}!${col}${row}`);

	section(sheet, 8, 'Revenue');
	line(sheet, I.retail, {
		label: 'Retail stores',
		unit: '$000',
		link: true,
		...all((col) => `=${revenueCell(col, R.retail)}`),
	});
	line(sheet, I.wholesale, {
		label: 'Wholesale',
		unit: '$000',
		link: true,
		...all((col) => `=${revenueCell(col, R.wholesale)}`),
	});
	line(sheet, I.online, {
		label: 'Online orders',
		unit: '$000',
		link: true,
		...all((col) => `=${revenueCell(col, R.online)}`),
	});
	line(sheet, I.revenue, {
		label: 'Total revenue',
		unit: '$000',
		kind: 'subtotal',
		base: sum(I.retail, I.online)('D'),
		cell: sum(I.retail, I.online),
	});
	sheet.groupRows(I.retail, I.online);
	line(sheet, I.cogs, {
		label: 'Cost of goods sold',
		unit: '$000',
		kind: 'label',
		link: true,
		...cost(K.cogs),
	});
	line(sheet, I.gross, {
		label: 'Gross profit',
		unit: '$000',
		kind: 'subtotal',
		...all((col) => `=${col}${I.revenue}+${col}${I.cogs}`),
	});
	line(sheet, I.grossMargin, {
		label: 'Gross margin',
		unit: '%',
		kind: 'ratio',
		fmt: FMT.percent,
		...all((col) => `=${col}${I.gross}/${col}${I.revenue}`),
	});

	section(sheet, 16, 'Operating costs');
	line(sheet, I.labour, { label: 'Store labour', unit: '$000', link: true, ...cost(K.labour) });
	line(sheet, I.rent, { label: 'Rent', unit: '$000', link: true, ...cost(K.rent) });
	line(sheet, I.utilities, {
		label: 'Utilities and other store costs',
		unit: '$000',
		link: true,
		...cost(K.utilities),
	});
	line(sheet, I.marketing, { label: 'Marketing', unit: '$000', link: true, ...cost(K.marketing) });
	line(sheet, I.headOffice, {
		label: 'Head office',
		unit: '$000',
		link: true,
		...cost(K.headOffice),
	});
	line(sheet, I.operating, {
		label: 'Total operating costs',
		unit: '$000',
		kind: 'subtotal',
		base: sum(I.labour, I.headOffice)('D'),
		cell: sum(I.labour, I.headOffice),
	});
	sheet.groupRows(I.labour, I.headOffice);
	line(sheet, I.ebitda, {
		label: 'EBITDA',
		unit: '$000',
		kind: 'total',
		...all((col) => `=${col}${I.gross}+${col}${I.operating}`),
	});
	line(sheet, I.ebitdaMargin, {
		label: 'EBITDA margin',
		unit: '%',
		kind: 'ratio',
		fmt: FMT.percent,
		...all((col) => `=${col}${I.ebitda}/${col}${I.revenue}`),
	});

	line(sheet, I.depreciation, {
		label: 'Depreciation',
		unit: '$000',
		kind: 'label',
		link: true,
		...cost(K.depreciation),
	});
	line(sheet, I.ebit, {
		label: 'Operating profit',
		unit: '$000',
		kind: 'subtotal',
		...all((col) => `=${col}${I.ebitda}+${col}${I.depreciation}`),
	});
	line(sheet, I.interest, {
		label: 'Interest on bank debt',
		unit: '$000',
		kind: 'label',
		// The base year is a typed actual; forecasts charge the rate on the opening balance.
		base: -196000,
		cell: (col, prev) => `=-${BALANCE}!${prev}${B.debt}*InterestRate`,
	});
	line(sheet, I.pretax, {
		label: 'Profit before tax',
		unit: '$000',
		kind: 'subtotal',
		...all((col) => `=${col}${I.ebit}+${col}${I.interest}`),
	});
	line(sheet, I.tax, {
		label: 'Tax',
		unit: '$000',
		kind: 'label',
		...all((col) => `=-MAX(0,${col}${I.pretax})*TaxRate`),
	});
	line(sheet, I.net, {
		label: 'Net income',
		unit: '$000',
		kind: 'total',
		...all((col) => `=${col}${I.pretax}+${col}${I.tax}`),
	});
	line(sheet, I.netMargin, {
		label: 'Net margin',
		unit: '%',
		kind: 'ratio',
		fmt: FMT.percent,
		...all((col) => `=${col}${I.net}/${col}${I.revenue}`),
	});
	sheet.heights({ [I.ebitda]: 19, [I.net]: 19 });
}
