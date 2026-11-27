/** The Orders sheet (an Excel table with calculated columns and a totals row) and the lookups. */
import type { Table } from '../../../../src/core/xlsx/model';
import type { Book, Look } from '../lib/book';
import { L, rangeOf } from '../lib/sheet';
import { type Order, PRODUCTS, TERRITORIES } from './data';
import { FMT, FOOTER, P, SHEETS, title } from './style';

/** The first data row of the Orders table (its header is the row above). */
export const FIRST = 6;

interface Column {
	name: string;
	width: number;
	fmt?: string;
	h?: NonNullable<Look['h']>;
	/** A text column that follows right-aligned figures stands off them by an indent. */
	indent?: boolean;
	/** A typed value of the order, or the calculated-column formula. */
	value?: (order: Order) => string | number;
	formula?: string;
	/** The totals row: a SUBTOTAL function number and its name in the table part. */
	total?: [number, string];
}

const row = (column: string) => `Orders[[#This Row],[${column}]]`;
const lookup = (column: string) =>
	`XLOOKUP(${row('Product')},Products[Product],Products[${column}])`;

const COLUMNS: Column[] = [
	{ name: 'Order ID', width: 13, value: (o) => o.id },
	{
		name: 'Order date',
		width: 12.5,
		fmt: FMT.date,
		h: 'right',
		value: (o) => o.date,
		total: [103, 'count'],
	},
	{
		name: 'Month',
		width: 10.5,
		fmt: FMT.monthYear,
		h: 'right',
		formula: `DATE(YEAR(${row('Order date')}),MONTH(${row('Order date')}),1)`,
	},
	{ name: 'Region', width: 17, indent: true, value: (o) => o.region },
	{ name: 'City', width: 16.5, value: (o) => o.city },
	{ name: 'Channel', width: 11, value: (o) => o.channel },
	{ name: 'Rep', width: 19, value: (o) => o.rep },
	{ name: 'Product', width: 27, value: (o) => o.product.name },
	{ name: 'Category', width: 18, formula: lookup('Category') },
	{ name: 'Units', width: 7.5, fmt: FMT.count, value: (o) => o.units, total: [109, 'sum'] },
	{ name: 'Unit price', width: 11, fmt: FMT.cents, formula: lookup('List price') },
	{
		name: 'Discount',
		width: 10,
		fmt: FMT.wholePercent,
		value: (o) => o.discount,
		total: [101, 'average'],
	},
	{
		name: 'Revenue',
		width: 13,
		fmt: FMT.cents,
		formula: `ROUND(${row('Units')}*${row('Unit price')}*(1-${row('Discount')}),2)`,
		total: [109, 'sum'],
	},
	{
		name: 'Cost',
		width: 12.5,
		fmt: FMT.cents,
		formula: `${row('Units')}*${lookup('Unit cost')}`,
		total: [109, 'sum'],
	},
	{
		name: 'Margin',
		width: 12.5,
		fmt: FMT.cents,
		formula: `${row('Revenue')}-${row('Cost')}`,
		total: [109, 'sum'],
	},
];

/** The cell range of the Orders table without its totals row, for the PivotTable source. */
export const ordersSource = (count: number): string =>
	`${SHEETS.orders}!B${FIRST - 1}:${L(COLUMNS.length)}${FIRST - 1 + count}`;

export function buildOrders(book: Book, list: Order[]): void {
	const sheet = book.sheet(SHEETS.orders);
	title(
		sheet,
		'Orders',
		`The 2025 order book: ${list.length} order lines. Category, price, revenue, cost and margin are calculated columns that look each product up in the Products table.`,
		P.lake,
	);
	sheet.ws.view.showGridLines = true;
	sheet.widths({ A: 2 });
	const header = FIRST - 1;
	const totals = FIRST + list.length;
	COLUMNS.forEach((column, i) => {
		const col = L(i + 1);
		sheet.widths({ [col]: column.width });
		const align = column.h ?? (column.fmt ? 'right' : 'left');
		sheet.set(`${col}${header}`, column.name, {
			bold: true,
			color: P.white,
			h: align,
			v: 'center',
			...(column.indent ? { indent: 1 } : {}),
		});
		const look: Look = {
			...(column.fmt ? { fmt: column.fmt } : {}),
			...(column.h ? { h: column.h } : {}),
			...(column.indent ? { indent: 1 } : {}),
		};
		list.forEach((order, n) => {
			const ref = `${col}${FIRST + n}`;
			if (column.formula) sheet.formula(ref, column.formula, look);
			else sheet.set(ref, column.value!(order), look);
		});
		const bold: Look = { ...look, bold: true };
		if (column.total)
			sheet.formula(`${col}${totals}`, `SUBTOTAL(${column.total[0]},Orders[${column.name}])`, {
				...bold,
				...(column.total[1] === 'count' ? { fmt: '#,##0" orders"' } : {}),
			});
		else sheet.set(`${col}${totals}`, i === 0 ? 'Total' : null, bold);
	});
	sheet.heights({ [header]: 20 });
	sheet.freeze(header, 2);
	const table: Table = {
		id: 1,
		name: 'Orders',
		displayName: 'Orders',
		range: rangeOf(`B${header}:${L(COLUMNS.length)}${totals}`),
		headerRow: true,
		totalsRow: true,
		columns: COLUMNS.map((column, i) => ({
			name: column.name,
			...(i === 0 ? { totalsRowLabel: 'Total' } : {}),
			...(column.total ? { totalsRowFunction: column.total[1] } : {}),
			...(column.formula ? { calculatedColumnFormula: column.formula } : {}),
		})),
		styleName: 'TableStyleMedium2',
		showRowStripes: true,
	};
	sheet.ws.tables.push(table);
	const last = totals - 1;
	// Deep discounts and the largest order lines stand out while scrolling.
	sheet.conditional(`M${FIRST}:M${last}`, {
		type: 'cellIs',
		operator: 'greaterThanOrEqual',
		formulas: ['0.3'],
		style: { font: { bold: true, color: { rgb: `FF${P.rust}` } } },
	});
	sheet.conditional(`N${FIRST}:N${last}`, {
		type: 'top10',
		rank: 10,
		style: {
			font: { bold: true, color: { rgb: `FF${P.pine}` } },
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFDCEBE5' } },
		},
	});
}

export function buildLookups(book: Book): void {
	const sheet = book.sheet(SHEETS.lookups);
	title(
		sheet,
		'Lookups',
		'Reference tables the order book reads: the product catalogue and the sales regions with their targets.',
		P.lichen,
	);
	sheet.ws.view.showGridLines = true;
	sheet.page(FOOTER, { tall: 1 });
	sheet.widths({ A: 2, B: 28, C: 19, D: 11, E: 11, F: 10, G: 3, H: 17, I: 20, J: 14, K: 40 });
	const head: Look = { bold: true, color: P.white, v: 'center' };
	['Product', 'Category', 'List price', 'Unit cost', 'Margin'].forEach((name, i) =>
		sheet.set(`${L(i + 1)}5`, name, { ...head, h: i > 1 ? 'right' : 'left' }),
	);
	const margin = '1-Products[[#This Row],[Unit cost]]/Products[[#This Row],[List price]]';
	PRODUCTS.forEach((product, n) => {
		const r = 6 + n;
		sheet.set(`B${r}`, product.name).set(`C${r}`, product.category);
		sheet
			.set(`D${r}`, product.price, { fmt: FMT.cents })
			.set(`E${r}`, product.cost, { fmt: FMT.cents });
		sheet.formula(`F${r}`, margin, { fmt: FMT.percent });
	});
	sheet.ws.tables.push({
		id: 2,
		name: 'Products',
		displayName: 'Products',
		range: rangeOf(`B5:F${5 + PRODUCTS.length}`),
		headerRow: true,
		totalsRow: false,
		columns: [
			{ name: 'Product' },
			{ name: 'Category' },
			{ name: 'List price' },
			{ name: 'Unit cost' },
			{ name: 'Margin', calculatedColumnFormula: margin },
		],
		styleName: 'TableStyleMedium5',
		showRowStripes: true,
	});
	sheet.conditional(`F6:F${5 + PRODUCTS.length}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'max' }],
		colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FFE6D9A6' }],
	});

	['Region', 'Regional manager', 'Annual target', 'Cities'].forEach((name, i) =>
		sheet.set(`${L(7 + i)}5`, name, { ...head, h: i === 2 ? 'right' : 'left' }),
	);
	TERRITORIES.forEach((territory, n) => {
		const r = 6 + n;
		sheet.set(`H${r}`, territory.region).set(`I${r}`, territory.manager);
		sheet.set(`J${r}`, territory.target, { fmt: FMT.money });
		sheet.set(`K${r}`, territory.cities.join(', '));
	});
	sheet.ws.tables.push({
		id: 3,
		name: 'Regions',
		displayName: 'Regions',
		range: rangeOf(`H5:K${5 + TERRITORIES.length}`),
		headerRow: true,
		totalsRow: false,
		columns: [
			{ name: 'Region' },
			{ name: 'Regional manager' },
			{ name: 'Annual target' },
			{ name: 'Cities' },
		],
		styleName: 'TableStyleMedium3',
		showRowStripes: true,
	});
	sheet.heights({ 5: 20 });
}
