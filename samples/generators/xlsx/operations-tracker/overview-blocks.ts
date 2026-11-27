/**
 * The spilling blocks of the Overview sheet. Each block is one dynamic array formula; the cells
 * below and beside its anchor are formatted to receive the spill.
 */
import { type Book, type Look, rgb } from '../lib/book';
import { DEPOTS } from './data';
import { FMT, HEAD, LINE, P, SHEETS, TOTAL, header, section } from './style';

const S = SHEETS.overview;
/** Anchor rows of the spilling blocks. */
export const AT = { lookup: 13, urgent: 17, depots: 32, split: 44, months: 55 } as const;
const URGENT_ROWS = 12;
export const ALL = 'All depots';

/** Formats the cells a spill will fill, column by column. */
function receive(
	book: Book,
	first: number,
	rows: number,
	columns: [string, Look][],
	line = LINE,
): void {
	const sheet = book.sheet(S);
	for (const [col, look] of columns)
		sheet.paint(`${col}${first}:${col}${first + rows - 1}`, { ...line, ...look });
}

export function blocks(book: Book): void {
	const sheet = book.sheet(S);
	const date: Look = { fmt: FMT.shortDate, h: 'right' };

	section(sheet, AT.lookup - 2, 'B', 'K', 'Work order lookup');
	header(sheet, AT.lookup - 1, [
		['B', 'Opened'],
		['C', 'Priority', 'left'],
		['D', 'Depot', 'left'],
		['E', 'Asset', 'left'],
		['F', 'Category', 'left'],
		['G', 'Assigned to', 'left'],
		['H', 'Status', 'left'],
		['I', 'Due'],
		['J', 'Closed'],
		['K', 'Days open'],
	]);
	// One XLOOKUP returns the whole record: a spill along the row.
	sheet.formula(
		`B${AT.lookup}`,
		'XLOOKUP(LookupOrder,WO_ID,HSTACK(WO_Opened,WO_Priority,WO_Depot,WO_Asset,WO_Category,WO_Tech,WO_Status,WO_Due,WO_Closed,WO_Age),"Not in the log")',
	);
	receive(book, AT.lookup, 1, [
		['B', date],
		['C', { indent: 1, bold: true }],
		['D', { indent: 1 }],
		['E', { indent: 1 }],
		['F', { indent: 1 }],
		['G', { indent: 1 }],
		['H', { indent: 1 }],
		['I', date],
		['J', { fmt: `${FMT.shortDate};;"open"`, h: 'right' }],
		['K', { fmt: '0' }],
	]);
	sheet.heights({ [AT.lookup]: 19 });

	section(sheet, AT.urgent - 2, 'B', 'K', 'Most urgent open work orders');
	header(sheet, AT.urgent - 1, [
		['B', '#'],
		['C', 'Work order', 'left'],
		['D', 'Depot', 'left'],
		['E', 'Asset', 'left'],
		['F', 'Priority', 'left'],
		['G', 'Assigned to', 'left'],
		['H', 'Status', 'left'],
		['I', 'Opened'],
		['J', 'Due'],
		['K', 'Overdue'],
	]);
	sheet.formula(`B${AT.urgent}`, `SEQUENCE(ROWS(C${AT.urgent}#))`);
	// Open work for the chosen depot, by priority and then by days overdue; the rank column is dropped.
	sheet.formula(
		`C${AT.urgent}`,
		`LET(rank,MATCH(WO_Priority,{"Critical";"High";"Normal";"Low"},0),` +
			'open,FILTER(HSTACK(WO_ID,WO_Depot,WO_Asset,WO_Priority,WO_Tech,WO_Status,WO_Opened,WO_Due,WO_Overdue,rank),' +
			`(WO_Status<>"Closed")*((Depot="${ALL}")+(WO_Depot=Depot))),` +
			`TAKE(DROP(SORT(open,{10,9},{1,-1}),,-1),${URGENT_ROWS}))`,
	);
	receive(book, AT.urgent, URGENT_ROWS, [
		['B', { color: P.muted, h: 'center' }],
		['C', { bold: true, indent: 1 }],
		['D', { indent: 1 }],
		['E', { indent: 1 }],
		['F', { indent: 1 }],
		['G', { indent: 1 }],
		['H', { indent: 1 }],
		['I', date],
		['J', date],
		['K', { fmt: FMT.overdue, bold: true, color: P.signal }],
	]);
	sheet.conditional(
		`F${AT.urgent}:F${AT.urgent + URGENT_ROWS - 1}`,
		{
			type: 'cellIs',
			operator: 'equal',
			formulas: ['"Critical"'],
			style: { font: { bold: true, color: rgb(P.signal) } },
		},
		{
			type: 'cellIs',
			operator: 'equal',
			formulas: ['"High"'],
			style: { font: { bold: true, color: rgb('B06A12') } },
		},
	);

	const d = AT.depots;
	const end = d + DEPOTS.length - 1;
	section(sheet, d - 2, 'B', 'K', 'Depot workload');
	// The depot name is wider than its column and runs on into the empty one beside it.
	header(sheet, d - 1, [
		['B', '#'],
		['C', 'Depot', 'left'],
		['E', 'Manager', 'left'],
		['F', 'Open'],
		['G', 'Overdue'],
		['H', 'Critical open'],
		['I', 'Closed'],
		['J', 'Days to close'],
		['K', 'Labour hours'],
	]);
	sheet.paint(`D${d - 1}`, HEAD);
	const depots = `C${d}#`;
	sheet.formula(`B${d}`, `SEQUENCE(ROWS(${depots}))`);
	sheet.formula(`C${d}`, 'SORT(UNIQUE(WO_Depot))');
	sheet.formula(`E${d}`, `XLOOKUP(${depots},Depots[Depot],Depots[Manager])`);
	sheet.formula(`F${d}`, `COUNTIFS(WO_Depot,${depots},WO_Status,"<>Closed")`);
	sheet.formula(`G${d}`, `COUNTIFS(WO_Depot,${depots},WO_Overdue,">0")`);
	sheet.formula(
		`H${d}`,
		`COUNTIFS(WO_Depot,${depots},WO_Priority,"Critical",WO_Status,"<>Closed")`,
	);
	sheet.formula(`I${d}`, `COUNTIFS(WO_Depot,${depots},WO_Status,"Closed")`);
	sheet.formula(`J${d}`, `AVERAGEIFS(WO_Age,WO_Depot,${depots},WO_Status,"Closed")`);
	sheet.formula(`K${d}`, `SUMIFS(WO_Hours,WO_Depot,${depots})`);
	receive(book, d, DEPOTS.length, [
		['B', { color: P.muted, h: 'center' }],
		['C', { bold: true, indent: 1 }],
		['D', {}],
		['E', { indent: 1 }],
		['F', { fmt: FMT.count }],
		['G', { fmt: FMT.count }],
		['H', { fmt: FMT.count }],
		['I', { fmt: FMT.count }],
		['J', { fmt: FMT.days }],
		['K', { fmt: FMT.count }],
	]);
	sheet
		.paint(`B${end + 1}:K${end + 1}`, TOTAL)
		.set(`C${end + 1}`, 'All depots', { ...TOTAL, indent: 1 });
	// Totals read the spills by reference, so they follow a depot being added.
	for (const col of ['F', 'G', 'H', 'I', 'K'])
		sheet.formula(`${col}${end + 1}`, `SUM(${col}${d}#)`, { ...TOTAL, fmt: FMT.count });
	sheet.formula(`J${end + 1}`, 'AVERAGEIFS(WO_Age,WO_Status,"Closed")', {
		...TOTAL,
		fmt: FMT.days,
	});
	sheet.conditional(`F${d}:F${end}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'max' },
		color: rgb('B9C7DD'),
	});
	sheet.conditional(`G${d}:G${end}`, {
		type: 'colorScale',
		thresholds: [{ type: 'num', value: '0' }, { type: 'max' }],
		colors: [rgb('FFFFFF'), rgb('F2C4BA')],
	});
	sheet.conditional(`J${d}:J${end}`, {
		type: 'colorScale',
		thresholds: [{ type: 'min' }, { type: 'max' }],
		colors: [rgb('FFFFFF'), rgb('F6D9A8')],
	});

	const t = AT.split;
	section(sheet, t - 2, 'B', 'E', 'Open work by priority');
	header(sheet, t - 1, [
		['C', 'Priority', 'left'],
		['D', 'Open'],
		['E', 'Share'],
	]);
	sheet.formula(
		`C${t}`,
		'LET(p,{"Critical";"High";"Normal";"Low"},HSTACK(p,COUNTIFS(WO_Priority,p,WO_Status,"<>Closed")))',
	);
	sheet.formula(`E${t}`, `D${t}:D${t + 3}/SUM(D${t}:D${t + 3})`);
	receive(book, t, 4, [
		['C', { bold: true, indent: 1 }],
		['D', { fmt: FMT.count }],
		['E', { fmt: FMT.percent }],
	]);
	sheet.conditional(`E${t}:E${t + 3}`, {
		type: 'dataBar',
		min: { type: 'num', value: '0' },
		max: { type: 'num', value: '1' },
		color: rgb('F3C98B'),
	});
	section(sheet, t - 2, 'G', 'K', 'Technicians with the most open work');
	header(sheet, t - 1, [
		['G', 'Technician', 'left'],
		['H', 'Open'],
		['I', 'Home depot', 'left'],
	]);
	sheet.paint(`J${t - 1}:K${t - 1}`, HEAD);
	sheet.formula(
		`G${t}`,
		'LET(who,UNIQUE(FILTER(WO_Tech,WO_Status<>"Closed")),n,COUNTIFS(WO_Tech,who,WO_Status,"<>Closed"),TAKE(SORTBY(HSTACK(who,n),n,-1,who,1),8))',
	);
	sheet.formula(`I${t}`, `XLOOKUP(G${t}:G${t + 7},WO_Tech,WO_Depot)`);
	receive(book, t, 8, [
		['G', { bold: true, indent: 1 }],
		['H', { fmt: FMT.count }],
		['I', { indent: 1 }],
		['J', {}],
		['K', {}],
	]);

	const m = AT.months;
	section(sheet, m - 2, 'B', 'K', 'Opened and closed by month');
	header(sheet, m - 1, [
		['C', 'Month', 'left'],
		['D', 'Opened'],
		['E', 'Closed'],
		['F', 'Net change'],
	]);
	const months = `C${m}#`;
	const next = `DATE(YEAR(${months}),MONTH(${months})+1,1)`;
	// The twelve months that end with the month of the As of date.
	sheet.formula(`C${m}`, 'DATE(YEAR(AsOf)-1,MONTH(AsOf)+SEQUENCE(12),1)');
	sheet.formula(`D${m}`, `COUNTIFS(WO_Opened,">="&${months},WO_Opened,"<"&${next})`);
	sheet.formula(`E${m}`, `COUNTIFS(WO_Closed,">="&${months},WO_Closed,"<"&${next})`);
	sheet.formula(`F${m}`, `D${m}#-E${m}#`);
	receive(book, m, 12, [
		['C', { fmt: FMT.month, h: 'left', indent: 1 }],
		['D', { fmt: FMT.count }],
		['E', { fmt: FMT.count }],
		['F', { fmt: '+0;-0;0', color: P.muted }],
	]);
}
