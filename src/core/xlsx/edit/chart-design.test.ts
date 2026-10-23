import { describe, expect, it } from 'vitest';
import { findChartStylePreset } from '../../chart/style-presets';
import { parseChartSpace } from '../../chart/parse-space';
import type { ChartObject } from '../model';
import { loadXlsx } from '../read';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write';
import {
	chartQuickLayoutPatch,
	chartStylePresetOf,
	chartStylePresetPatch,
	chartTakesQuickLayout,
} from './chart-design';
import { createEditSession } from './session';
import JSZip from 'jszip';

const column = (): Omit<ChartObject, 'kind'> => ({
	chartType: 'column',
	anchor: { from: { row: 0, col: 3, colOffset: 0, rowOffset: 0 }, ext: { cx: 4e6, cy: 3e6 } },
	showLegend: true,
	series: [
		{ name: 'Sales', categories: ['a', 'b'], values: [2, 4] },
		{ name: 'Cost', categories: ['a', 'b'], values: [1, 3] },
	],
});

async function chartPart(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return zip.file('xl/charts/chart1.xml')!.async('string');
}

describe('xlsx Chart Styles presets', () => {
	it('pins the preset colours, recognises them and saves them', async () => {
		const book = createWorkbook();
		const session = createEditSession(book);
		session.addChart(0, column());
		const chart = () => book.sheets[0]!.drawings[0] as ChartObject;
		session.updateChart(0, 0, chartStylePresetPatch(chart(), 'pastel')!);
		expect(chartStylePresetOf(chart())).toBe('pastel');
		const loaded = await loadXlsx(await saveXlsx(book));
		const back = loaded.sheets[0]!.drawings[0] as ChartObject;
		expect(chartStylePresetOf(back)).toBe('pastel');
		const colors = findChartStylePreset('pastel')!.colors();
		expect(`#${back.series[1]!.drawingColor?.value}`.toLowerCase()).toBe(colors[1]!.toLowerCase());
		expect(chartStylePresetPatch(chart(), 'nope')).toBeUndefined();
	});

	it('pins a pie chart per point', () => {
		const pie = { ...column(), kind: 'chart', chartType: 'pie' } as ChartObject;
		pie.series = [pie.series[0]!];
		const patch = chartStylePresetPatch(pie, 'colorful')!;
		expect(Object.keys(patch.series![0]!.pointColors!)).toEqual(['0', '1']);
		expect(chartStylePresetOf({ ...pie, ...patch } as ChartObject)).toBe('colorful');
	});
});

describe('xlsx Quick Layout', () => {
	it('applies Layout 5 to the summary and writes the rest into a new chart part', async () => {
		const book = createWorkbook();
		const session = createEditSession(book);
		session.addChart(0, column());
		const chart = () => book.sheets[0]!.drawings[0] as ChartObject;
		session.updateChart(0, 0, chartQuickLayoutPatch(chart(), 5)!);
		expect(chart()).toMatchObject({ quickLayout: 5, title: 'Chart Title', showLegend: false });
		expect(chart().axes?.find((axis) => axis.kind === 'val')?.title?.text).toBe('Axis Title');
		const xml = await chartPart(await saveXlsx(book));
		const space = parseChartSpace(xml).chartSpace;
		expect(space.plotArea.dataTable).toMatchObject({ showKeys: true });
		expect(space.plotArea.axes.find((axis) => axis.kind === 'val')?.title?.text).toBe('Axis Title');
		expect(space.legend).toBeUndefined();
		session.undo();
		expect(chart().quickLayout).toBeUndefined();
	});

	it('applies Layout 2 to a kept chart part, keeping the source detail', async () => {
		const first = createWorkbook();
		createEditSession(first).addChart(0, column());
		const zip = await JSZip.loadAsync(await saveXlsx(first));
		const source = await zip.file('xl/charts/chart1.xml')!.async('string');
		zip.file(
			'xl/charts/chart1.xml',
			source.replace('</c:catAx>', '<c:extLst><c:ext uri="retained"/></c:extLst></c:catAx>'),
		);
		const book = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const kept = book.sheets[0]!.drawings[0] as ChartObject;
		expect(kept.partName).toBeDefined();
		const patch = chartQuickLayoutPatch(kept, 2)!;
		expect(patch.series?.[0]?.dataLabels?.chart?.showValue).toBe(true);
		createEditSession(book).updateChart(0, 0, patch);
		const xml = await chartPart(await saveXlsx(book));
		expect(xml).toContain('uri="retained"');
		const space = parseChartSpace(xml).chartSpace;
		expect(space.legend?.position).toBe('t');
		expect(space.plotArea.groups[0]).toMatchObject({ overlap: -25 });
		expect(space.plotArea.groups[0]?.dataLabels).toMatchObject({
			showValue: true,
			position: 'outEnd',
		});
		expect(space.plotArea.axes.find((axis) => axis.kind === 'val')?.deleted).toBe(true);
	});

	it('refuses chart families without layouts', () => {
		const pie = { ...column(), kind: 'chart', chartType: 'pie' } as ChartObject;
		expect(chartTakesQuickLayout(pie)).toBe(false);
		expect(chartQuickLayoutPatch(pie, 1)).toBeUndefined();
		expect(chartQuickLayoutPatch({ ...pie, chartType: 'line' }, 12)).toBeUndefined();
	});
});
