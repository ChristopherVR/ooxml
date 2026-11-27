/** The Costs sheet: cost of goods sold, store costs, overheads, capital expenditure, depreciation. */
import type { Book } from '../lib/book';
import { A } from './assumptions';
import { R, revenueCell } from './revenue';
import { P, SHEETS, frame, line, section, yearHeader } from './style';

const S = SHEETS.costs;
const ASSUME = SHEETS.assumptions;

export const K = {
	ingredients: 9,
	packaging: 10,
	cogs: 11,
	labour: 14,
	rent: 15,
	utilities: 16,
	storeCosts: 17,
	marketing: 20,
	headOffice: 21,
	overheads: 22,
	operating: 24,
	fitOut: 27,
	maintenance: 28,
	extension: 29,
	capex: 30,
	existingDepreciation: 33,
	newDepreciation: 34,
	depreciation: 35,
} as const;

export function buildCosts(book: Book): void {
	const sheet = book.sheet(S);
	frame(
		sheet,
		'Costs',
		'Cost of goods sold, store operating costs, overheads and the capital programme. Dollars, shown in thousands.',
		P.sage,
	);
	yearHeader(sheet, 'Line item', true);
	sheet.widths({ A: 2, B: 40, C: 10, D: 13, 'E:I': 12.5 }).freeze(5, 3);
	const total = (col: string) => revenueCell(col, R.total);
	const retail = (col: string) => revenueCell(col, R.retail);
	// Years since the base year, from the header row, so inflation compounds without a helper row.
	const elapsed = (col: string) => `(${col}$5-$D$5)`;
	const stores = (col: string) =>
		`AVERAGE(${revenueCell(col, R.opening)},${revenueCell(col, R.closing)})`;
	const sum = (from: number, to: number) => (col: string) => `=SUM(${col}${from}:${col}${to})`;

	section(sheet, 8, 'Cost of goods sold');
	line(sheet, K.ingredients, {
		label: 'Ingredients',
		unit: '$000',
		base: `=${total('D')}*${ASSUME}!D${A.ingredients}`,
		cell: (col) => `=${total(col)}*${ASSUME}!${col}${A.ingredients}`,
	});
	line(sheet, K.packaging, {
		label: 'Packaging',
		unit: '$000',
		base: `=${total('D')}*PackagingPct`,
		cell: (col) => `=${total(col)}*PackagingPct`,
	});
	line(sheet, K.cogs, {
		label: 'Cost of goods sold',
		unit: '$000',
		kind: 'subtotal',
		base: sum(K.ingredients, K.packaging)('D'),
		cell: sum(K.ingredients, K.packaging),
	});
	sheet.groupRows(K.ingredients, K.packaging);

	section(sheet, 13, 'Store operating costs');
	line(sheet, K.labour, {
		label: 'Store labour',
		unit: '$000',
		base: `=${retail('D')}*${ASSUME}!D${A.labour}`,
		cell: (col) => `=${retail(col)}*${ASSUME}!${col}${A.labour}`,
	});
	line(sheet, K.rent, {
		label: 'Rent',
		unit: '$000',
		base: `=${revenueCell('D', R.closing)}*RentPerStore`,
		cell: (col) => `=${stores(col)}*RentPerStore*(1+RentInflation)^${elapsed(col)}`,
	});
	line(sheet, K.utilities, {
		label: 'Utilities and other store costs',
		unit: '$000',
		base: `=${revenueCell('D', R.closing)}*UtilitiesPerStore`,
		cell: (col) => `=${stores(col)}*UtilitiesPerStore*(1+UtilitiesInflation)^${elapsed(col)}`,
	});
	line(sheet, K.storeCosts, {
		label: 'Store operating costs',
		unit: '$000',
		kind: 'subtotal',
		base: sum(K.labour, K.utilities)('D'),
		cell: sum(K.labour, K.utilities),
	});
	sheet.groupRows(K.labour, K.utilities);

	section(sheet, 19, 'Overheads');
	line(sheet, K.marketing, {
		label: 'Marketing',
		unit: '$000',
		base: `=${total('D')}*MarketingPct`,
		cell: (col) => `=${total(col)}*MarketingPct`,
	});
	line(sheet, K.headOffice, {
		label: 'Head office',
		unit: '$000',
		base: '=HeadOfficeBase',
		cell: (col) => `=HeadOfficeBase*(1+HeadOfficeGrowth)^${elapsed(col)}`,
	});
	line(sheet, K.overheads, {
		label: 'Overheads',
		unit: '$000',
		kind: 'subtotal',
		base: sum(K.marketing, K.headOffice)('D'),
		cell: sum(K.marketing, K.headOffice),
	});
	sheet.groupRows(K.marketing, K.headOffice);
	line(sheet, K.operating, {
		label: 'Total operating costs',
		unit: '$000',
		kind: 'total',
		base: `=D${K.storeCosts}+D${K.overheads}`,
		cell: (col) => `=${col}${K.storeCosts}+${col}${K.overheads}`,
	});

	section(sheet, 26, 'Capital expenditure');
	line(sheet, K.fitOut, {
		label: 'New store fit-out',
		unit: '$000',
		cell: (col) => `=${revenueCell(col, R.newStores)}*CapexPerStore`,
	});
	line(sheet, K.maintenance, {
		label: 'Maintenance of the estate',
		unit: '$000',
		cell: (col) => `=${total(col)}*MaintenanceCapexPct`,
	});
	line(sheet, K.extension, {
		label: 'Central bakery extension',
		unit: '$000',
		link: true,
		cell: (col) => `=${ASSUME}!${col}${A.extension}`,
	});
	line(sheet, K.capex, {
		label: 'Capital expenditure',
		unit: '$000',
		kind: 'subtotal',
		cell: sum(K.fitOut, K.extension),
	});
	sheet.groupRows(K.fitOut, K.extension);

	section(sheet, 32, 'Depreciation');
	line(sheet, K.existingDepreciation, {
		label: 'Assets held at the base year',
		unit: '$000',
		base: '=ExistingDepreciation',
		cell: () => '=ExistingDepreciation',
	});
	line(sheet, K.newDepreciation, {
		label: 'Assets added in the plan',
		unit: '$000',
		// Straight line from the year of spending: each year adds a slice of every year so far.
		cell: (col) => `=SUM($E${K.capex}:${col}${K.capex})/UsefulLife`,
	});
	line(sheet, K.depreciation, {
		label: 'Depreciation',
		unit: '$000',
		kind: 'subtotal',
		base: sum(K.existingDepreciation, K.newDepreciation)('D'),
		cell: sum(K.existingDepreciation, K.newDepreciation),
	});
	sheet.groupRows(K.existingDepreciation, K.newDepreciation);
}
