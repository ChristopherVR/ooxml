import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read';
import { saveXlsx } from './index';
import { chartXml } from './chart';
import { patchChartPart } from './chart-patch';
import { parseChart } from '../read/chart';
import type { ChartObject } from '../model';

function chart(
	chartType: ChartObject['chartType'],
	axisVisible: boolean,
	labelsVisible: boolean,
): ChartObject {
	return {
		kind: 'chart',
		chartType,
		showLegend: false,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [{ categories: ['A', 'B'], values: [10, 20] }],
		formatting: {
			sourceXml: '',
			entries: {
				categoryAxis: { sourceXml: '', axisVisible, labelsVisible },
				valueAxis: { sourceXml: '', axisVisible, labelsVisible },
			},
		},
	};
}

describe('XLSX axis formatting export', () => {
	it.each(['column', 'bar', 'line', 'area', 'scatter', 'radar'] as const)(
		'round trips axis flags for new %s charts through repeated saves',
		async (type) => {
			for (const [axisVisible, labelsVisible] of [
				[true, true],
				[true, false],
				[false, false],
			] as const) {
				let book = createWorkbook();
				createEditSession(book).addChart(0, chart(type, axisVisible, labelsVisible));
				for (let save = 0; save < 2; save++) {
					book = await loadXlsx(await saveXlsx(book));
					const back = book.sheets[0]!.drawings[0] as ChartObject;
					for (const part of ['categoryAxis', 'valueAxis'] as const)
						expect(back.formatting?.entries[part]).toMatchObject({
							axisVisible,
							labelsVisible,
						});
				}
			}
		},
	);

	it('patches flags on kept charts and retains unknown axis XML', () => {
		const model = chart('column', true, true);
		const source = chartXml(model)
			.replace('<c:tickLblPos val="nextTo"/>', '<c:tickLblPos val="high"/>')
			.replace('</c:catAx>', '<c:extLst><c:ext uri="retained"/></c:extLst></c:catAx>');
		const imported = parseChart(source, model.anchor, 'xl/charts/chart1.xml');
		expect(patchChartPart(source, imported)).toBeUndefined();
		imported.formatting!.entries.categoryAxis!.labelsVisible = false;
		const patched = patchChartPart(source, imported)!;
		expect(patched).toContain('uri="retained"');
		expect(
			parseChart(patched, model.anchor, '').formatting?.entries.categoryAxis?.labelsVisible,
		).toBe(false);
	});

	it.each(['type', 'grouping', 'series'] as const)(
		'retains flags when %s changes regenerate a kept chart',
		(change) => {
			const model = chart('column', true, false),
				source = chartXml(model);
			if (change === 'type') model.chartType = 'line';
			if (change === 'grouping') model.grouping = 'stacked';
			if (change === 'series') model.series.push({ categories: ['A', 'B'], values: [30, 40] });
			const back = parseChart(patchChartPart(source, model)!, model.anchor, '');
			expect(back.formatting?.entries.categoryAxis?.labelsVisible).toBe(false);
			expect(back.formatting?.entries.valueAxis?.labelsVisible).toBe(false);
		},
	);
});
