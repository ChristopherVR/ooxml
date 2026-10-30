import { describe, expect, it } from 'vitest';
import { miniPackage, roundTrip, ws } from './mini-package-fixtures';

const ROWS = '<row r="1"><c r="A1"><v>1</v></c></row>';
const VIEWS = (pane: string) =>
	`<sheetViews><sheetView workbookViewId="0" topLeftCell="B2">${pane}</sheetView></sheetViews>`;
const SPLIT =
	'<pane xSplit="3000" ySplit="1500" topLeftCell="D9" activePane="bottomRight" state="split"/>';
const book = (views: string) => miniPackage({ sheets: [{ name: 'S', xml: ws(ROWS, views) }] });

describe('split panes', () => {
	it('reads xSplit and ySplit as twentieths of a point', async () => {
		const { wb } = await roundTrip(await book(VIEWS(SPLIT)));
		const view = wb.sheets[0]!.view;
		expect(view.freeze).toBeUndefined();
		expect(view.split).toEqual({
			x: 3000,
			y: 1500,
			topLeft: { row: 1, col: 1 },
			activePane: 'bottomRight',
		});
		expect(view.topLeft).toEqual({ row: 8, col: 3 });
	});

	it('keeps an opened split verbatim on an unedited save', async () => {
		const { part } = await roundTrip(await book(VIEWS(SPLIT)));
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain(SPLIT);
	});

	it('keeps the split when the selection is edited', async () => {
		const { part, again } = await roundTrip(await book(VIEWS(SPLIT)), (wb) => {
			wb.sheets[0]!.view.selection = {
				active: { row: 9, col: 4 },
				ranges: [{ start: { row: 9, col: 4 }, end: { row: 9, col: 4 } }],
			};
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain('state="split"');
		expect(xml).toContain('xSplit="3000"');
		expect(xml).toContain('ySplit="1500"');
		expect(xml).toContain('topLeftCell="B2"');
		expect(xml).not.toContain('state="frozen"');
		expect(again.sheets[0]?.view.split).toEqual({
			x: 3000,
			y: 1500,
			topLeft: { row: 1, col: 1 },
			activePane: 'bottomRight',
		});
		expect(again.sheets[0]?.view.topLeft).toEqual({ row: 8, col: 3 });
	});

	it('writes a one-axis split with the matching active pane', async () => {
		const { part, again } = await roundTrip(await book(VIEWS('')), (wb) => {
			const view = wb.sheets[0]!.view;
			delete view.topLeft;
			delete view.split;
			view.split = { x: 0, y: 2400 };
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain('<pane ySplit="2400" activePane="bottomLeft" state="split"/>');
		expect(again.sheets[0]?.view.split).toEqual({ x: 0, y: 2400, activePane: 'bottomLeft' });
	});
});
