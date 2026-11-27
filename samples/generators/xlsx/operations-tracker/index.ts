/** `operations-tracker.xlsx`: the maintenance work-order tracker of Vale Orrin Logistics. */
import type { SampleDefinition } from '../../lib/types';
import { calculate, finish } from '../lib/finish';
import { ORDER_COUNT, workOrders } from './data';
import { buildDepots } from './depots';
import { HEADER, buildLog } from './log';
import { buildOverview, decorateOverview } from './overview';
import { buildReport, printReport } from './report';
import { SHEETS, createBook } from './style';

async function generate(): Promise<Uint8Array> {
	const book = createBook();
	const orders = workOrders();
	buildOverview(book, orders);
	buildLog(book, orders);
	buildReport(book);
	buildDepots(book);
	const session = calculate(book);
	decorateOverview(book, session);
	printReport(book, session);
	// The header row of the log repeats on every printed page.
	session.setPrintTitles(book.sheet(SHEETS.log).index, {
		rows: { start: HEADER - 1, end: HEADER - 1 },
	});
	return finish(book, session, {
		title: 'Vale Orrin Logistics operations tracker',
		subject: 'Maintenance work orders',
		description:
			'A work-order log with a live view built from dynamic array formulas, comments, protection and a print setup. Vale Orrin Logistics is fictitious.',
		keywords:
			'operations, work orders, dynamic arrays, FILTER, XLOOKUP, comments, protection, print',
		company: 'Vale Orrin Logistics',
	});
}

export const operationsTracker: SampleDefinition = {
	id: 'operations-tracker',
	product: 'xlsx',
	fileName: 'operations-tracker.xlsx',
	title: 'Operations tracker',
	description: `A maintenance tracker for Vale Orrin Logistics: a log of ${ORDER_COUNT} work orders, a live overview that spills from dynamic array formulas, a printable depot report, notes and threaded comments, and protected sheets with unlocked inputs.`,
	features: [
		'Log of 4,000 rows with an AutoFilter and a stored sort state',
		'Dynamic arrays that spill: FILTER, SORT, SORTBY, UNIQUE, XLOOKUP, LET, SEQUENCE, TAKE, DROP, HSTACK, TEXTJOIN',
		'Spill references (A1#) in totals and COUNTIFS, and a chart over spilled ranges',
		'Legacy notes and threaded comments with replies and people',
		'Sheet protection without a password, with unlocked input cells',
		'Print setup: print area, repeated titles, landscape, scaling, headers and footers, page breaks',
		'Conditional formatting for overdue and priority rows, data bars, icon sets',
		'Data validation: list, date and custom rules',
		'Internal, web and mail hyperlinks; an Excel table; named ranges',
	],
	generate,
};
