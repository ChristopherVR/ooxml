// Insert Chart for Excel's extended families: a chart over the selection, its data read from the
// sheet, edited before the first save and saved as a new chartEx part with its style parts.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	applyChartExSpaceDataSheetEdits,
	CHART_EX_FAMILIES,
	chartExFamilyOf,
	type ChartExFamily,
} from '../../chart';
import { getCell } from '../cells';
import { loadXlsx } from '../read';
import { saveXlsx } from '../write';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { xlsxChartExDataSheet } from './chart-ex';
import { buildChartEx } from '../ui/dialogs/insert-chart-ex-model';
import type { Workbook } from '../model';

function book(): Workbook {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	const rows: (string | number)[][] = [
		['Group', 'Cause', 'Count'],
		['Shipping', 'Late', 42],
		['Shipping', 'Damaged', 18],
		['Order', 'Wrong item', 27],
		['Order', 'Missing', 9],
	];
	rows.forEach((row, r) => row.forEach((value, c) => session.setCellValue(0, r, c, value)));
	return workbook;
}

const RANGE = { start: { row: 0, col: 1 }, end: { row: 4, col: 2 } };

async function inserted(family: ChartExFamily, range = RANGE) {
	const workbook = book();
	const session = createEditSession(workbook);
	const index = session.addChartEx(0, buildChartEx(workbook, 0, range, family, { title: family }));
	return { workbook, session, index };
}

function charts(workbook: Workbook) {
	return workbook.sheets[0]!.drawings.flatMap((d, i) =>
		d.kind === 'unsupported' && d.chartExSpace ? [{ d, i }] : [],
	);
}

describe('inserting an Excel extended chart', () => {
	it.each(CHART_EX_FAMILIES)(
		'saves and reloads a new %s chart reading the sheet',
		async (family) => {
			const { workbook } = await inserted(family);
			const saved = await saveXlsx(workbook);
			const zip = await JSZip.loadAsync(saved);
			const types = await zip.file('[Content_Types].xml')!.async('string');
			expect(types).toContain(
				'PartName="/xl/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"',
			);
			expect(types).toContain('application/vnd.ms-office.chartstyle+xml');
			expect(types).toContain('application/vnd.ms-office.chartcolorstyle+xml');
			const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
			expect(drawing).toContain(
				'<mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">',
			);
			expect(await zip.file('xl/charts/chartEx1.xml')!.async('string')).toContain(
				'<cx:f>Sheet1!$C$2:$C$5</cx:f>',
			);
			const back = await loadXlsx(saved);
			const [chart] = charts(back);
			expect(chart).toBeDefined();
			expect(chartExFamilyOf(chart!.d.chartExSpace!)).toBe(family);
			expect(xlsxChartExDataSheet(back, 0, chart!.i).series[0]!.values).toEqual([42, 18, 27, 9]);
		},
	);

	it('reads leading text columns as treemap levels', async () => {
		const { workbook, index } = await inserted('treemap', {
			start: { row: 0, col: 0 },
			end: { row: 4, col: 2 },
		});
		const sheet = xlsxChartExDataSheet(workbook, 0, index);
		expect(sheet.categoryLevels).toEqual([
			['Late', 'Damaged', 'Wrong item', 'Missing'],
			['Shipping', 'Shipping', 'Order', 'Order'],
		]);
		const saved = await JSZip.loadAsync(await saveXlsx(workbook));
		expect(await saved.file('xl/charts/chartEx1.xml')!.async('string')).toContain(
			'<cx:f>Sheet1!$A$2:$B$5</cx:f>',
		);
	});

	it('edits the data of a chart not yet saved in the sheet cells, as one undo step', async () => {
		const { workbook, session, index } = await inserted('pareto');
		session.editChartEx(0, index, (space) =>
			applyChartExSpaceDataSheetEdits(space, [
				{ kind: 'value', series: 0, category: 3, value: 50 },
			]),
		);
		expect(getCell(workbook.sheets[0]!, 4, 2)?.value).toBe(50);
		session.undo();
		expect(getCell(workbook.sheets[0]!, 4, 2)?.value).toBe(9);
		session.undo();
		expect(charts(workbook)).toHaveLength(0);
	});
});
