/**
 * The Explorer sheet: a PivotTable over the order book with a report filter, connected to two
 * slicers and a timeline. All four are created through the core's edit session.
 */
import type { EditSession } from '../../../../src/core/xlsx/edit/types';
import type { DrawingAnchor } from '../../../../src/core/xlsx/model';
import type { Book } from '../lib/book';
import { anchor } from '../lib/chart';
import { ordersSource } from './orders';
import { FMT, FOOTER, P, SHEETS, section, title } from './style';

const PIVOT = 'RevenueExplorer';

/** Moves the slicer or timeline frame called `name` to a new anchor. */
function place(
	session: EditSession,
	book: Book,
	sheet: number,
	name: string,
	at: DrawingAnchor,
): void {
	const index = book.wb.sheets[sheet]!.drawings.findIndex(
		(drawing) =>
			(drawing.kind === 'slicer' || drawing.kind === 'timeline') && drawing.name === name,
	);
	if (index < 0) throw new Error(`No frame named ${name}`);
	session.setDrawingAnchor(sheet, index, at);
}

export function buildExplorer(book: Book, session: EditSession, count: number): void {
	const sheet = book.sheet(SHEETS.explorer);
	const s = sheet.index;
	title(
		sheet,
		'Revenue explorer',
		'A PivotTable of revenue by category, product and channel. Filter it with the Rep report filter, the slicers and the timeline.',
		P.ember,
	);
	sheet.page(FOOTER, { portrait: true, tall: 1 });
	sheet.widths({ A: 2, B: 30, 'C:F': 13, G: 3, 'H:I': 12, J: 2, 'K:L': 12 });
	section(sheet, 5, 'B', 'F', 'Order date');
	section(sheet, 5, 'H', 'L', 'Slice by');
	section(sheet, 14, 'B', 'F', 'Revenue by category, product and channel');

	session.createPivotTable(
		s,
		{ row: 15, col: 1 },
		{
			source: ordersSource(count),
			name: PIVOT,
			layout: 'compact',
			rowField: 'Category',
			secondRowField: 'Product',
			columnField: 'Channel',
			valueField: 'Revenue',
			pageFilter: { field: 'Rep' },
		},
	);
	session.updatePivotTable(s, PIVOT, {
		display: {
			style: {
				name: 'PivotStyleMedium2',
				showRowHeaders: true,
				showColHeaders: true,
				showRowStripes: true,
			},
		},
	});
	const pivot = sheet.ws.pivotTables!.find((entry) => entry.name === PIVOT)!;
	// Whole dollars in the value area; the PivotTable keeps cell formats when it is refreshed.
	session.applyStyle(
		s,
		[
			{
				start: { row: pivot.location.start.row + 1, col: pivot.location.start.col + 1 },
				end: pivot.location.end,
			},
		],
		{ numFmt: FMT.money },
	);

	const region = session.insertSlicer(s, { kind: 'pivot', pivot: PIVOT, field: 'Region' });
	const channel = session.insertSlicer(s, { kind: 'pivot', pivot: PIVOT, field: 'Channel' });
	const timeline = session.insertTimeline(s, { pivot: PIVOT, field: 'Order date' });
	session.updateSlicer(region, { caption: 'Region', style: 'SlicerStyleLight1' });
	session.updateSlicer(channel, { caption: 'Channel', style: 'SlicerStyleLight2' });
	session.updateTimeline(timeline, {
		caption: 'Order date',
		level: 2,
		style: 'TimeSlicerStyleLight1',
	});
	place(session, book, s, timeline, anchor('B6', 'G13'));
	place(session, book, s, region, anchor('H6', 'J17'));
	place(session, book, s, channel, anchor('K6', 'M13'));
}
