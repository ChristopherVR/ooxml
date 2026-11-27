/**
 * The Work orders sheet: the log itself, filtered and sorted, with rules for overdue and
 * priority rows, notes, threaded comments and a print setup.
 */
import type { Book, Look } from '../lib/book';
import { L, rangeOf } from '../lib/sheet';
import { PEOPLE } from '../../lib/fiction';
import { SLA, type WorkOrder } from './data';
import { FMT, HEAD, P, SHEETS } from './style';

const S = SHEETS.log;
/** The header row and the first data row. */
export const HEADER = 4;
export const FIRST = 5;

interface Column {
	name: string;
	/** The workbook name of the column's data range. */
	range: string;
	width: number;
	look?: Look;
	/** A text column that carries a look (text is left-aligned, numbers and dates right). */
	text?: boolean;
	value?: (order: WorkOrder) => string | number | null;
	formula?: (row: number) => string;
}

const date: Look = { fmt: FMT.date, h: 'right' };
const COLUMNS: Column[] = [
	{ name: 'Work order', range: 'WO_ID', width: 12, value: (o) => o.id },
	{ name: 'Opened', range: 'WO_Opened', width: 12, look: date, value: (o) => o.opened },
	{
		name: 'Depot',
		range: 'WO_Depot',
		width: 24,
		look: { indent: 1 },
		text: true,
		value: (o) => o.depot,
	},
	{ name: 'Asset', range: 'WO_Asset', width: 21, value: (o) => o.asset },
	{ name: 'Category', range: 'WO_Category', width: 20, value: (o) => o.category },
	{ name: 'Priority', range: 'WO_Priority', width: 10, value: (o) => o.priority },
	{ name: 'Assigned to', range: 'WO_Tech', width: 19, value: (o) => o.technician },
	{ name: 'Status', range: 'WO_Status', width: 17, value: (o) => o.status },
	{ name: 'Due', range: 'WO_Due', width: 12, look: date, value: (o) => o.due },
	{ name: 'Closed', range: 'WO_Closed', width: 12, look: date, value: (o) => o.closed ?? null },
	{ name: 'Hours', range: 'WO_Hours', width: 8, look: { fmt: '0.00' }, value: (o) => o.hours },
	{
		name: 'Parts cost',
		range: 'WO_Cost',
		width: 11,
		look: { fmt: FMT.money },
		value: (o) => o.cost,
	},
	{
		name: 'Days open',
		range: 'WO_Age',
		width: 10,
		look: { fmt: '0' },
		formula: (r) => `IF(J${r}="",AsOf,J${r})-B${r}`,
	},
	{
		name: 'Overdue',
		range: 'WO_Overdue',
		width: 9,
		look: { fmt: FMT.overdue },
		formula: (r) => `IF(AND(H${r}<>"Closed",I${r}<AsOf),AsOf-I${r},0)`,
	},
];
const LAST_COL = L(COLUMNS.length - 1);

/** Notes on the column headers: what the column means. */
const NOTES: [string, string][] = [
	[
		'Priority',
		`Critical: the asset is out of service, due in ${SLA.Critical} days.\nHigh: a defect that limits use, ${SLA.High} days.\nNormal: ${SLA.Normal} days.\nLow: ${SLA.Low} days.`,
	],
	['Status', 'Open, In progress, Waiting for parts or Closed. Only closed work has a Closed date.'],
	[
		'Days open',
		'Days from Opened to Closed, or to the As of date on the Overview sheet while the work is not closed.',
	],
	[
		'Overdue',
		'Days past the Due date for work that is not closed, as of the date on the Overview sheet. Blank when on time.',
	],
];

function conversations(book: Book, orders: WorkOrder[]): void {
	const sheet = book.sheet(S);
	const rowOf = (test: (order: WorkOrder) => boolean): number => {
		const index = orders.findIndex(test);
		if (index < 0) throw new Error('No work order matches a comment');
		return FIRST + index;
	};
	const used = new Set<number>();
	const pick = (test: (order: WorkOrder) => boolean): number => {
		const row = rowOf((order) => test(order) && !used.has(FIRST + orders.indexOf(order)));
		used.add(row);
		return row;
	};
	const day = (n: number, time: string) => `2026-01-${String(n).padStart(2, '0')}T${time}:00.00`;
	const [dispatcher, planner, lead, buyer] = [PEOPLE[0], PEOPLE[3], PEOPLE[9], PEOPLE[5]];
	const critical = pick((o) => o.priority === 'Critical' && o.status !== 'Closed');
	sheet.thread(
		`H${critical}`,
		{
			author: dispatcher,
			text: 'This unit is booked on the Marrow Bay night trunk tomorrow. Can it be back on the road by 18:00?',
			date: day(14, '08:12'),
		},
		[
			{
				author: lead,
				text: 'Parts arrived this morning. Fitting now, road test at 15:00.',
				date: day(14, '09:40'),
			},
			{
				author: dispatcher,
				text: 'Thanks. I will hold the booking and confirm after the road test.',
				date: day(14, '09:52'),
			},
		],
	);
	const waiting = pick((o) => o.status === 'Waiting for parts' && o.opened < 45960);
	sheet.thread(
		`H${waiting}`,
		{
			author: planner,
			text: 'This has waited for a part since last year. Is the order still live with the supplier?',
			date: day(12, '14:05'),
		},
		[
			{
				author: buyer,
				text: 'The supplier discontinued the part. An equivalent is on order and due on 22 January.',
				date: day(13, '10:21'),
			},
			{
				author: planner,
				text: 'Understood. I have moved the asset to the reserve pool until then.',
				date: day(13, '11:02'),
			},
		],
	);
	const costly = pick((o) => o.status === 'Closed' && o.cost > 2300);
	sheet.thread(
		`L${costly}`,
		{
			author: planner,
			text: 'Parts cost is well above the usual for this category. Was the gearbox replaced rather than repaired?',
			date: day(9, '16:30'),
		},
		[
			{
				author: lead,
				text: 'Replaced under the exchange scheme. The core credit will come through as a separate line.',
				date: day(10, '08:15'),
			},
		],
		true,
	);
	const overdue = pick((o) => o.status === 'Open' && o.due < 46030);
	sheet.thread(`I${overdue}`, {
		author: dispatcher,
		text: 'Due date has passed and nobody has picked this up. Reassigning at the morning stand-up.',
		date: day(15, '07:45'),
	});
}

export function buildLog(book: Book, orders: WorkOrder[]): void {
	const sheet = book.sheet(S);
	const last = FIRST + orders.length - 1;
	sheet.present(P.blue, { grid: true });
	sheet.set('A1', 'Work order log', { font: 'display', size: 16, bold: true, color: P.navy });
	sheet.set(
		'A2',
		`${String(orders.length).replace(/\B(?=(\d{3})+$)/g, ',')} maintenance work orders since 2 January 2025, sorted with overdue work first and then newest first. Overdue rows are shaded; filter and sort from the header row.`,
		{ color: P.muted, size: 9, italic: true },
	);
	sheet.set(`${LAST_COL}1`, 'Back to the overview', { color: P.blue, underline: true, h: 'right' });
	sheet.linkTo(`${LAST_COL}1`, `${SHEETS.overview}!A1`, 'Go to the Overview sheet');
	sheet.heights({ 1: 27, 2: 16, 3: 6, [HEADER]: 22 });
	COLUMNS.forEach((column, i) => {
		const col = L(i);
		sheet.widths({ [col]: column.width });
		sheet.set(`${col}${HEADER}`, column.name, {
			...HEAD,
			h: column.look && !column.text ? 'right' : 'left',
			...(column.text ? { indent: 1 } : {}),
		});
		orders.forEach((order, n) => {
			const ref = `${col}${FIRST + n}`;
			if (column.formula) sheet.formula(ref, column.formula(FIRST + n), column.look);
			else sheet.set(ref, column.value!(order), column.look);
		});
		book.wb.definedNames.push({
			name: column.range,
			formula: `'${S}'!$${col}$${FIRST}:$${col}$${last}`,
		});
	});
	book.wb.definedNames.push({ name: 'WO_Log', formula: `'${S}'!$A$${FIRST}:$${LAST_COL}$${last}` });
	sheet.freeze(HEADER, 1);

	const data = rangeOf(`A${FIRST}:${LAST_COL}${last}`);
	sheet.ws.autoFilter = {
		range: rangeOf(`A${HEADER}:${LAST_COL}${last}`),
		// The sort the rows are stored in: most overdue first, then newest first.
		sortState: {
			range: data,
			conditions: [
				{ ref: rangeOf(`N${FIRST}:N${last}`), descending: true },
				{ ref: rangeOf(`B${FIRST}:B${last}`), descending: true },
			],
		},
	};

	const body = `A${FIRST}:${LAST_COL}${last}`;
	sheet.conditional(body, {
		type: 'expression',
		formula: `$N${FIRST}>0`,
		style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFBE9E5' } } },
	});
	sheet.conditional(
		`F${FIRST}:F${last}`,
		{
			type: 'cellIs',
			operator: 'equal',
			formulas: ['"Critical"'],
			style: { font: { bold: true, color: { rgb: `FF${P.signal}` } } },
		},
		{
			type: 'cellIs',
			operator: 'equal',
			formulas: ['"High"'],
			style: { font: { color: { rgb: 'FFB06A12' } } },
		},
	);
	sheet.conditional(`H${FIRST}:H${last}`, {
		type: 'containsText',
		text: 'Waiting',
		style: { font: { italic: true, color: { rgb: 'FFB06A12' } } },
	});
	sheet.conditional(`N${FIRST}:N${last}`, {
		type: 'cellIs',
		operator: 'greaterThan',
		formulas: ['0'],
		style: { font: { bold: true, color: { rgb: `FF${P.signal}` } } },
	});
	sheet.conditional(`L${FIRST}:L${last}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: { rgb: 'FFB9C7DD' },
	});

	for (const [name, text] of NOTES)
		sheet.note(
			`${L(COLUMNS.findIndex((column) => column.name === name))}${HEADER}`,
			PEOPLE[3],
			text,
		);
	conversations(book, orders);

	sheet.ws.pageSetup = {
		orientation: 'landscape',
		paperSize: 9,
		fitToWidth: 1,
		fitToHeight: 0,
		printArea: rangeOf(`A1:${LAST_COL}${last}`),
		margins: { left: 0.4, right: 0.4, top: 0.7, bottom: 0.6, header: 0.3, footer: 0.3 },
		header: '&L&BVale Orrin Logistics&B&CWork order log&RAs of 15 January 2026',
		footer: '&LInternal: maintenance planning&CPage &P of &N&R&A',
	};
	sheet.ws.printOptions = { horizontalCentered: true };
}
