import { describe, expect, it } from 'vitest';
import { miniPackage, roundTrip, ws } from './mini-package-fixtures';

const ROWS = '<row r="1"><c r="A1"><v>1</v></c></row>';
const SPLIT =
	'<pane xSplit="3000" ySplit="1500" topLeftCell="D9" activePane="bottomRight" state="split"/>';
const SELECTIONS =
	'<selection pane="topRight" activeCell="E1" sqref="E1"/>' +
	'<selection pane="bottomLeft" activeCell="A9" sqref="A9:B9"/>' +
	'<selection pane="bottomRight" activeCell="D9" sqref="D9"/>';
const doc = () =>
	miniPackage({
		sheets: [
			{
				name: 'S',
				xml: ws(
					ROWS,
					`<sheetViews><sheetView workbookViewId="0" topLeftCell="B2">${SPLIT}${SELECTIONS}</sheetView></sheetViews>`,
				),
			},
		],
	});

describe('split pane selections and scroll', () => {
	it('reads the selection of every pane', async () => {
		const { wb } = await roundTrip(await doc());
		const view = wb.sheets[0]!.view;
		expect(view.selection?.active).toEqual({ row: 8, col: 3 });
		expect(view.split?.paneSelections?.topRight?.active).toEqual({ row: 0, col: 4 });
		expect(view.split?.paneSelections?.bottomLeft?.ranges).toEqual([
			{ start: { row: 8, col: 0 }, end: { row: 8, col: 1 } },
		]);
	});

	it('keeps every pane selection and writes the current first cells', async () => {
		const { part, again } = await roundTrip(await doc(), (wb) => {
			wb.sheets[0]!.view.topLeft = { row: 20, col: 6 };
			wb.sheets[0]!.view.split!.topLeft = { row: 4, col: 2 };
		});
		const xml = (await part('xl/worksheets/sheet1.xml')) ?? '';
		expect(xml).toContain('topLeftCell="G21"');
		expect(xml).toContain('<sheetView');
		expect(xml).toContain('topLeftCell="C5"');
		expect(xml).toContain('<selection pane="topRight" activeCell="E1" sqref="E1"/>');
		expect(xml).toContain('<selection pane="bottomLeft" activeCell="A9" sqref="A9:B9"/>');
		expect(xml).toContain('<selection pane="bottomRight" activeCell="D9" sqref="D9"/>');
		expect(again.sheets[0]?.view.topLeft).toEqual({ row: 20, col: 6 });
		expect(again.sheets[0]?.view.split?.topLeft).toEqual({ row: 4, col: 2 });
		expect(again.sheets[0]?.view.split?.paneSelections?.topRight).toBeDefined();
	});
});
