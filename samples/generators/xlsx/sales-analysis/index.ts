/** `sales-analysis.xlsx`: the 2025 order book of Kestrelline Outfitters, analysed. */
import type { SampleDefinition } from '../../lib/types';
import { calculate, finish } from '../lib/finish';
import { buildAnalysis } from './analysis';
import { decorateDashboard } from './charts';
import { buildDashboard } from './dashboard';
import { orders } from './data';
import { buildExplorer, numberFallbackShapes } from './explorer';
import { FIRST, buildLookups, buildOrders } from './orders';
import { SHEETS, createBook } from './style';

async function generate(): Promise<Uint8Array> {
	const book = createBook();
	const list = orders();
	buildOrders(book, list);
	buildLookups(book);
	buildAnalysis(book, list);
	buildDashboard(book, list.length);
	// The PivotTable caches calculated columns, so the workbook is calculated before it is built.
	const session = calculate(book);
	buildExplorer(book, session, list.length);
	decorateDashboard(book, session);
	// The table header repeats on every printed page of the order book.
	session.setPrintTitles(book.sheet(SHEETS.orders).index, {
		rows: { start: FIRST - 2, end: FIRST - 2 },
	});
	return finish(
		book,
		session,
		{
			title: 'Kestrelline Outfitters sales analysis 2025',
			subject: 'Sales analysis',
			description:
				'An order book as an Excel table, analysed with a PivotTable, slicers, a timeline, charts and sparklines. Kestrelline Outfitters is fictitious.',
			keywords: 'sales, table, PivotTable, slicer, timeline, charts, sparklines',
			company: 'Kestrelline Outfitters',
		},
		numberFallbackShapes,
	);
}

export const salesAnalysis: SampleDefinition = {
	id: 'sales-analysis',
	product: 'xlsx',
	fileName: 'sales-analysis.xlsx',
	title: 'Sales analysis',
	description:
		'A year of orders for Kestrelline Outfitters as an Excel table with calculated columns, explored through a PivotTable with slicers and a timeline, and summarised on a dashboard of charts and sparklines.',
	features: [
		'Excel table of 960 rows with calculated columns, structured references and a totals row',
		'Lookup tables read with XLOOKUP',
		'PivotTable with row, column and value fields, a report filter and its cache',
		'Two slicers and a timeline connected to the PivotTable',
		'Combination chart: columns with a line on a secondary axis',
		'Doughnut, bar and stacked area charts',
		'Line and column sparklines with markers',
		'Conditional formatting: data bars, colour scales, icon sets, top and bottom rules',
		'A designed dashboard on a custom theme',
	],
	generate,
};
