/**
 * The Overview sheet: a live view of the log built from dynamic array formulas. Each block is one
 * formula that spills; the cells below and beside the anchors are formatted to receive it.
 */
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import { type Book, type Look, rgb } from '../lib/book';
import { anchor, chartObject } from '../lib/chart';
import { type Tile, tiles as drawTiles } from '../lib/tiles';
import { DEPOTS, type WorkOrder } from './data';
import { ALL, AT, blocks } from './overview-blocks';
import { AS_OF, FMT, FOOTER, INPUT, P, SHEETS } from './style';

const S = SHEETS.overview;
/** `d` is the depot criterion: every depot (a wildcard) or the one picked. */
const scoped = (body: string) => `LET(d,IF(Depot="${ALL}","*",Depot),${body})`;

function controls(book: Book, sample: string): void {
	const sheet = book.sheet(S);
	sheet.paint('B5:K5', { fill: P.paper, v: 'center' });
	const label: Look = { fill: P.paper, color: P.muted, h: 'right', v: 'center' };
	sheet.set('C5', 'As of', label).set('D5', AS_OF, { ...INPUT, fmt: FMT.date, h: 'center' });
	sheet.set('E5', 'Depot', label).set('F5', ALL, { ...INPUT, h: 'center' });
	sheet.set('H5', 'Look up work order', label).set('I5', sample, { ...INPUT, h: 'center' });
	sheet.heights({ 4: 6, 5: 24, 6: 6 });
	sheet.validate('D5', {
		type: 'date',
		operator: 'between',
		formula1: '45658',
		formula2: '46387',
		showInputMessage: true,
		promptTitle: 'As of date',
		prompt: 'Overdue days and open ages are measured to this date.',
		showErrorMessage: true,
		errorStyle: 'stop',
		errorTitle: 'Date out of range',
		error: 'Enter a date in 2025 or 2026.',
	});
	sheet.validate('F5', {
		type: 'list',
		formula1: `"${[ALL, ...DEPOTS.map((depot) => depot.name)].join(',')}"`,
		showDropDown: true,
		showInputMessage: true,
		promptTitle: 'Depot',
		prompt: 'Pick one depot, or All depots.',
		showErrorMessage: true,
		errorStyle: 'stop',
		errorTitle: 'Unknown depot',
		error: 'Choose a depot from the list.',
	});
	sheet.validate('I5', {
		type: 'custom',
		formula1: 'COUNTIF(WO_ID,I5)=1',
		showInputMessage: true,
		promptTitle: 'Work order',
		prompt: 'Type a work order number, for example WO-23990.',
		showErrorMessage: true,
		errorStyle: 'warning',
		errorTitle: 'Not in the log',
		error: 'No work order has this number. Keep it anyway?',
	});
	sheet.note(
		'D5',
		'Planning',
		"The tracker does not use today's date: change this cell to report as of another day.",
	);
	book.wb.definedNames.push(
		{ name: 'AsOf', formula: `${S}!$D$5`, comment: 'The date the tracker reports as of' },
		{ name: 'Depot', formula: `${S}!$F$5` },
		{ name: 'LookupOrder', formula: `${S}!$I$5` },
	);
}

function tiles(book: Book): void {
	const sheet = book.sheet(S);
	const open = 'WO_Status,"<>Closed"';
	const recent = 'WO_Closed,">="&(AsOf-30),WO_Closed,"<="&AsOf';
	const specs: Tile[] = [
		{
			first: 'B',
			last: 'D',
			label: 'Open work orders',
			value: scoped(`COUNTIFS(WO_Depot,d,${open})`),
			note: scoped(`COUNTIFS(WO_Depot,d,WO_Status,"Waiting for parts")&" waiting for parts"`),
			fmt: FMT.count,
			color: P.navy,
		},
		{
			first: 'E',
			last: 'F',
			label: 'Overdue',
			value: scoped('COUNTIFS(WO_Depot,d,WO_Overdue,">0")'),
			note: scoped('"Oldest is "&MAXIFS(WO_Overdue,WO_Depot,d)&" days past due"'),
			fmt: FMT.count,
			color: P.signal,
		},
		{
			first: 'G',
			last: 'H',
			label: 'Critical and open',
			value: scoped(`COUNTIFS(WO_Depot,d,WO_Priority,"Critical",${open})`),
			note: scoped(`COUNTIFS(WO_Depot,d,WO_Priority,"High",${open})&" more at high priority"`),
			fmt: FMT.count,
			color: P.amber,
		},
		{
			first: 'I',
			last: 'K',
			label: 'Closed in the last 30 days',
			value: scoped(`COUNTIFS(WO_Depot,d,${recent})`),
			note: scoped(
				`"Average "&TEXT(AVERAGEIFS(WO_Age,WO_Depot,d,${recent}),"0.0")&" days to close"`,
			),
			fmt: FMT.count,
			color: P.teal,
		},
	];
	drawTiles(sheet, 7, specs, { fill: P.paper, accent: P.navy, muted: P.muted });
	// Where critical work is late, by name; a reassuring line when there is none.
	sheet.formula(
		'B10',
		'LET(late,SORT(UNIQUE(FILTER(WO_Depot,(WO_Priority="Critical")*(WO_Overdue>0),""))),IF(INDEX(late,1)="","No critical work is overdue.","Critical work is overdue at: "&TEXTJOIN(", ",TRUE,late)))',
		{ color: P.signal, size: 9, bold: true, v: 'center', indent: 1 },
	);
	sheet.conditional('B10', {
		type: 'beginsWith',
		text: 'No critical',
		style: { font: { bold: true, color: rgb(P.teal) } },
	});
	sheet.heights({ 10: 20 });
}

export function buildOverview(book: Book, orders: WorkOrder[]): void {
	const sheet = book.sheet(S);
	sheet.present(P.navy).page(FOOTER, { portrait: true, tall: 1 });
	sheet.widths({ A: 2, B: 9, C: 13, D: 24, E: 22, F: 19, G: 21, H: 18, I: 13, J: 13, K: 12 });
	sheet.heights({ 1: 9, 2: 30, 3: 17 });
	sheet.set('B2', 'Vale Orrin Logistics', { font: 'display', size: 20, bold: true, color: P.navy });
	sheet.set(
		'B3',
		'Maintenance work orders: a live view of the log, built from dynamic array formulas. Change a shaded cell and every block follows.',
		{ color: P.muted },
	);
	// An open, in-progress job makes a good first lookup.
	controls(book, orders.find((order) => order.status === 'In progress')!.id);
	tiles(book);
	blocks(book);
	const foot = AT.months + 14;
	const link: Look = { color: P.blue, underline: true };
	sheet
		.set(`B${foot}`, 'Open the full log', link)
		.linkTo(`B${foot}`, `'${SHEETS.log}'!A4`, 'Go to the work order log');
	sheet
		.set(`D${foot}`, 'Depot report', link)
		.linkTo(`D${foot}`, `'${SHEETS.report}'!A1`, 'Go to the printable depot report');
	sheet
		.set(`E${foot}`, 'Depots', link)
		.linkTo(`E${foot}`, `${SHEETS.depots}!A1`, 'Go to the depot list');
	sheet.set(`F${foot}`, 'Maintenance handbook', link);
	sheet.linkOut(
		`F${foot}`,
		'https://handbook.valeorrin.example.invalid/maintenance/work-orders',
		'How work orders are raised, prioritised and closed',
	);
	sheet.set(
		`B${foot + 2}`,
		'This sheet is protected without a password so the formulas cannot be overwritten; the three shaded cells stay editable. Vale Orrin Logistics is fictitious.',
		{ italic: true, color: P.muted, size: 9 },
	);
	sheet.ws.protection = { sheet: true, allow: ['formatColumns', 'formatRows'] };
	book.wb.activeSheet = sheet.index;
}

/** The chart beside the monthly block, drawn from the spilled ranges once they hold values. */
export function decorateOverview(book: Book, session: EditSession): void {
	const sheet = book.sheet(S);
	const m = AT.months;
	const col = (letter: string) => `${S}!$${letter}$${m}:$${letter}$${m + 11}`;
	const series = (letter: string, color: string) => ({
		name: `${S}!$${letter}$${m - 1}`,
		categories: col('C'),
		values: col(letter),
		color,
	});
	session.addChart(
		sheet.index,
		chartObject(
			book.wb,
			'Opened and closed by month',
			'Columns of work orders opened each month with a line of work orders closed, for the last twelve months.',
			anchor(`G${m - 1}`, `L${m + 12}`),
			{
				ink: P.ink,
				grid: P.rule,
				legend: 'b',
				categoryFormat: 'mmm',
				valueFormat: '#,##0',
				groups: [
					{ type: 'column', gapWidth: 60, series: [series('D', P.blue)] },
					{ type: 'line', series: [{ ...series('E', P.amber), markers: true }] },
				],
			},
		),
	);
}
