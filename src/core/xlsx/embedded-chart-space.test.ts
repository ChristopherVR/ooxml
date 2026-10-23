import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { applyChartDataSheetEditsToSpace, chartSpaceDataSheet } from '../chart/data-sheet-space';
import { applyChartEditIntent } from '../chart/edit/intent';
import type { ChartObject } from './model';
import { readEmbeddedChartSpace, writeEmbeddedChartSpace } from './embedded-chart-space';
import { getCell } from './cells';
import { loadXlsx } from './read';
import { gridChartWorkbook, readEmbeddedChart } from './ui/dialogs/grid-chart';

const ROWS = [
	[null, 'Plan', 'Actual'],
	['North', 4, 2],
	['South', 3, 5],
];

describe('embedded chart as a ChartSpace', () => {
	it('reads the chart and its data sheet', async () => {
		const bytes = (await gridChartWorkbook(ROWS, 'column'))!;
		const found = (await readEmbeddedChartSpace(bytes))!;
		expect(found.partName).toMatch(/^xl\/charts\/chart\d+\.xml$/);
		expect(chartSpaceDataSheet(found.space)).toEqual({
			categories: ['North', 'South'],
			series: [
				{ name: 'Plan', values: [4, 3] },
				{ name: 'Actual', values: [2, 5] },
			],
		});
		expect(await readEmbeddedChartSpace(new Uint8Array([1, 2, 3]))).toBeUndefined();
	});

	it('writes data edits into the cells and the series references', async () => {
		const bytes = (await gridChartWorkbook(ROWS, 'column'))!;
		const { space } = (await readEmbeddedChartSpace(bytes))!;
		const edited = applyChartDataSheetEditsToSpace(space, [
			{ kind: 'value', series: 1, category: 0, value: 9 },
			{ kind: 'add-category' },
			{ kind: 'series-name', series: 0, name: 'Budget' },
			{ kind: 'add-series' },
		]);
		const out = await writeEmbeddedChartSpace(bytes, edited);
		const book = await loadXlsx(out);
		const sheet = book.sheets[0]!;
		expect(getCell(sheet, 0, 1)?.value).toBe('Budget');
		expect(getCell(sheet, 1, 2)?.value).toBe(9);
		expect(getCell(sheet, 3, 0)?.value).toBe('Cat 3');
		expect(getCell(sheet, 0, 3)?.value).toBe('Series 3');
		const chart = sheet.drawings.find((d): d is ChartObject => d.kind === 'chart')!;
		expect(chart.series.map((series) => series.valuesRef)).toEqual([
			'Sheet1!$B$2:$B$4',
			'Sheet1!$C$2:$C$4',
			'Sheet1!$D$2:$D$4',
		]);
		expect(chart.series[0]?.nameRef).toBe('Sheet1!$B$1');
		const view = (await readEmbeddedChart(out))!;
		expect(view.view.series.map((series) => series.name)).toEqual(['Budget', 'Actual', 'Series 3']);
	});

	it('writes formatting edits without touching the cells', async () => {
		const bytes = (await gridChartWorkbook(ROWS, 'column'))!;
		const { space, partName } = (await readEmbeddedChartSpace(bytes))!;
		const edited = applyChartEditIntent(space, {
			op: 'fill',
			target: { kind: 'series', seriesIndex: 0 },
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
		});
		const out = await writeEmbeddedChartSpace(
			bytes,
			applyChartEditIntent(edited, { op: 'quickLayout', layout: 2 }),
		);
		const back = (await readEmbeddedChartSpace(out))!;
		expect(back.partName).toBe(partName);
		const zip = await JSZip.loadAsync(out);
		expect(Object.values(zip.files).some((entry) => entry.dir)).toBe(false);
		const fill = back.space.plotArea.groups[0]!.series[0]!.spPr?.fill;
		expect(fill).toMatchObject({ kind: 'solid', color: { value: 'FF0000' } });
		expect(back.space.legend?.position).toBe('t');
		expect(getCell((await loadXlsx(out)).sheets[0]!, 1, 1)?.value).toBe(4);
	});
});
