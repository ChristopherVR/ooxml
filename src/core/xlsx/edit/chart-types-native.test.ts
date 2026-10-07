import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-types.json';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import type { ChartObject } from '../model';
import { buildChart } from '../ui/dialogs/insert-chart-model';
import { createEditSession } from './session';

async function nativeWorkbook(xml: string) {
	const workbook = createWorkbook();
	workbook.sheets[0]!.name = 'Chart data';
	const session = createEditSession(workbook);
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales', 'Cost'],
		['Jan', 10, 4],
		['Feb', 20, 8],
		['Mar', 30, 9],
	]);
	session.addChart(
		0,
		buildChart(workbook, 0, { start: { row: 0, col: 0 }, end: { row: 3, col: 2 } }, 'column'),
	);
	const zip = await JSZip.loadAsync(await saveXlsx(workbook));
	zip.file('xl/charts/chart1.xml', xml);
	return loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
}

describe(`Excel ${native.excelVersion} build ${native.excelBuild} chart type/grouping transitions`, () => {
	for (const source of native.cases) {
		it(`reads and changes native ${source.type}/${source.grouping}`, async () => {
			const workbook = await nativeWorkbook(source.xml);
			const chart = () => workbook.sheets[0]!.drawings[0] as ChartObject;
			expect(chart()).toMatchObject({
				chartType: source.type,
				grouping: source.grouping,
				title: source.title,
			});
			expect(
				chart().series.map((series) => ({ name: series.name, values: series.values })),
			).toEqual(source.series.map(({ name, values }) => ({ name, values })));
			const before = structuredClone(chart());
			const session = createEditSession(workbook);
			for (const target of native.cases) {
				session.updateChart(0, 0, {
					chartType: target.type as ChartObject['chartType'],
					grouping: target.grouping as NonNullable<ChartObject['grouping']>,
				});
				const loaded = await loadXlsx(await saveXlsx(workbook));
				const after = loaded.sheets[0]!.drawings[0] as ChartObject;
				expect(after).toMatchObject({
					chartType: target.type,
					grouping: target.grouping,
					title: target.title,
				});
				expect(after.series).toMatchObject(before.series);
				session.undo();
				expect(chart()).toEqual(before);
			}
		});
	}
});
