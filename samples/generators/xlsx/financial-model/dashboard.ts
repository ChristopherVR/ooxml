/** The Dashboard: headline tiles, the five-year summary with conditional formats, two charts. */
import { type Book, type Look, rgb } from '../lib/book';
import { type Tile, tiles as drawTiles } from '../lib/tiles';
import { K } from './costs';
import { R, revenueCell } from './revenue';
import { B, F, I } from './rows';
import { FMT, FOOTER, LOOK, P, SHEETS, YEARS, section } from './style';

const S = SHEETS.dashboard;
const INCOME = `'${SHEETS.income}'`;
const BALANCE = `'${SHEETS.balance}'`;
const CASH = `'${SHEETS.cash}'`;
/** The model columns the summary shows: the base year and the five forecast years. */
const COLUMNS = ['D', ...YEARS];
export const FIRST = 11;

interface Metric {
	label: string;
	unit: string;
	fmt: string;
	/** The formula of the model column `col`; undefined leaves the base year empty. */
	cell: (col: string) => string | undefined;
	detail?: boolean;
	bold?: boolean;
	/** Amounts get a compound growth rate; ratios do not. */
	growth?: boolean;
}

const METRICS: Metric[] = [
	{
		label: 'Total revenue',
		unit: '$000',
		fmt: FMT.thousands,
		bold: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.revenue}`,
	},
	{
		label: 'Retail stores',
		unit: '$000',
		fmt: FMT.thousands,
		detail: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.retail}`,
	},
	{
		label: 'Wholesale',
		unit: '$000',
		fmt: FMT.thousands,
		detail: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.wholesale}`,
	},
	{
		label: 'Online orders',
		unit: '$000',
		fmt: FMT.thousands,
		detail: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.online}`,
	},
	{
		label: 'Revenue growth',
		unit: '%',
		fmt: FMT.percent,
		cell: (c) => (c === 'D' ? undefined : revenueCell(c, R.growth)),
	},
	{
		label: 'Gross margin',
		unit: '%',
		fmt: FMT.percent,
		cell: (c) => `${INCOME}!${c}${I.grossMargin}`,
	},
	{
		label: 'EBITDA',
		unit: '$000',
		fmt: FMT.thousands,
		bold: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.ebitda}`,
	},
	{
		label: 'EBITDA margin',
		unit: '%',
		fmt: FMT.percent,
		cell: (c) => `${INCOME}!${c}${I.ebitdaMargin}`,
	},
	{
		label: 'Net income',
		unit: '$000',
		fmt: FMT.thousands,
		bold: true,
		growth: true,
		cell: (c) => `${INCOME}!${c}${I.net}`,
	},
	{
		label: 'Cash at year end',
		unit: '$000',
		fmt: FMT.thousands,
		growth: true,
		cell: (c) => (c === 'D' ? `${BALANCE}!D${B.cash}` : `${CASH}!${c}${F.closing}`),
	},
	{ label: 'Bank debt', unit: '$000', fmt: FMT.thousands, cell: (c) => `${BALANCE}!${c}${B.debt}` },
	{
		label: 'Capital expenditure',
		unit: '$000',
		fmt: FMT.thousands,
		cell: (c) => (c === 'D' ? undefined : `${SHEETS.costs}!${c}${K.capex}`),
	},
	{
		label: 'Bakeries at year end',
		unit: 'stores',
		fmt: FMT.count,
		growth: true,
		cell: (c) => revenueCell(c, R.closing),
	},
];
export const rowOf = (label: string): number => FIRST + METRICS.findIndex((m) => m.label === label);
export const LAST = FIRST + METRICS.length - 1;

function tiles(book: Book): void {
	const final = revenueCell('I', R.closing);
	const specs: Tile[] = [
		{
			first: 'B',
			last: 'C',
			label: '="Revenue in FY"&(StartYear+4)',
			value: 'INDEX(Revenue_Total,5)',
			fmt: FMT.millions,
			note: `"Compound growth of "&TEXT((INDEX(Revenue_Total,5)/${INCOME}!D${I.revenue})^(1/5)-1,"0.0%")&" a year"`,
		},
		{
			first: 'D',
			last: 'F',
			label: '="EBITDA margin in FY"&(StartYear+4)',
			value: 'INDEX(EBITDA,5)/INDEX(Revenue_Total,5)',
			fmt: '0.0%',
			note: `"From "&TEXT(${INCOME}!D${I.ebitdaMargin},"0.0%")&" in the base year"`,
		},
		{
			first: 'G',
			last: 'I',
			label: '="Cash at the end of FY"&(StartYear+4)',
			value: 'INDEX(Closing_Cash,5)',
			fmt: FMT.millions,
			note: '"Lowest year end: $"&TEXT(MIN(Closing_Cash)/1000000,"0.0")&"m"',
		},
		{
			first: 'J',
			last: 'K',
			label: '="Bakeries at the end of FY"&(StartYear+4)',
			value: final,
			fmt: '0',
			note: `(${final}-OpeningStores)&" more than in the base year"`,
		},
	];
	drawTiles(book.sheet(S), 5, specs, { fill: P.cream, accent: P.crust, muted: P.muted });
}

export function buildDashboard(book: Book): void {
	const sheet = book.sheet(S);
	sheet.present(P.teal).page(FOOTER, { tall: 1 });
	sheet.widths({ A: 2, B: 25, C: 9, 'D:I': 12.5, J: 14, K: 20 });
	sheet.heights({ 1: 9, 2: 30, 3: 17, 4: 9 });
	sheet.set('B2', 'Brindlemoor Bakehouse', { named: 'Title' });
	sheet.formula(
		'B3',
		'"Five-year plan dashboard.  Scenario: "&Scenario&".  FY"&StartYear&" to FY"&(StartYear+4)&", dollars in thousands unless stated."',
		{ color: P.muted },
	);
	sheet.set('K3', 'Change the scenario', { color: P.inputInk, underline: true, h: 'right' });
	sheet.linkTo('K3', `${SHEETS.assumptions}!D8`, 'Go to the scenario selector');
	tiles(book);

	section(sheet, 9, 'Five-year summary', 'K');
	const head: Look = { bold: true, color: P.white, fill: P.ink, v: 'center' };
	sheet.paint('B10:K10', head);
	sheet.set('B10', 'Metric', { ...head, indent: 1 }).set('C10', 'Unit', head);
	COLUMNS.forEach((col, i) =>
		sheet.formula(`${col}10`, `${SHEETS.assumptions}!${col}5`, {
			...head,
			h: 'right',
			fmt: i === 0 ? FMT.actualYear : FMT.forecastYear,
		}),
	);
	sheet
		.set('J10', 'Trend', { ...head, h: 'center' })
		.set('K10', 'Growth a year', { ...head, h: 'right' });
	METRICS.forEach((metric, i) => {
		const row = FIRST + i;
		const base: Look = { bottom: P.sand, v: 'center', ...(metric.bold ? { bold: true } : {}) };
		sheet.paint(`B${row}:K${row}`, base);
		sheet.set(`B${row}`, metric.label, { ...base, indent: metric.detail ? 2 : 1 });
		sheet.set(`C${row}`, metric.unit, { ...base, ...LOOK.unit, bold: false });
		for (const col of COLUMNS) {
			const formula = metric.cell(col);
			if (formula) sheet.formula(`${col}${row}`, formula, { ...base, fmt: metric.fmt });
		}
		if (metric.growth)
			sheet.formula(`K${row}`, `(I${row}/D${row})^(1/5)-1`, {
				...base,
				fmt: '+0.0%;-0.0%;0.0%',
				bold: false,
				color: P.muted,
			});
		sheet.heights({ [row]: 19 });
	});
	sheet.groupRows(rowOf('Retail stores'), rowOf('Online orders'));

	const revenue = rowOf('Total revenue');
	const growth = rowOf('Revenue growth');
	const margin = rowOf('EBITDA margin');
	const cash = rowOf('Cash at year end');
	sheet.conditional(`D${revenue}:I${revenue}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: rgb('E3B79F'),
	});
	sheet.conditional(`D${cash}:I${cash}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: rgb('A9C4C9'),
	});
	sheet.conditional(`E${growth}:I${growth}`, {
		type: 'iconSet',
		iconSet: '3Arrows',
		thresholds: [
			{ type: 'percent', value: '0' },
			{ type: 'num', value: '0.1' },
			{ type: 'num', value: '0.15' },
		],
	});
	sheet.conditional(`D${margin}:I${margin}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
		colors: [rgb('FBF6EE'), rgb('DDE5CF'), rgb('A9BC8C')],
	});
	section(sheet, LAST + 2, 'Revenue, margin and funding', 'K');
}
