import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read';
import { saveXlsx } from './index';
import { chartXml } from './chart';
import { parseChart } from '../read/chart';
import type { ChartObject } from '../model';
import type { ChartPatch } from '../edit/charts';
import { withChartPart } from './test-support/chart-parity-cases';

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

const savedPart = async (book: Parameters<typeof saveXlsx>[0]) =>
	(await JSZip.loadAsync(await saveXlsx(book))).file('xl/charts/chart1.xml')!.async('string');

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

	it('edits axis options on kept charts and retains unknown axis XML', async () => {
		const model = chart('column', true, true);
		const source = chartXml(model)
			.replace('<c:tickLblPos val="nextTo"/>', '<c:tickLblPos val="high"/>')
			.replace('</c:catAx>', '<c:extLst><c:ext uri="retained"/></c:extLst></c:catAx>');
		const book = await withChartPart(source);
		expect(await savedPart(book)).toBe(source);
		createEditSession(book).editChart(0, 0, {
			op: 'axis',
			axisIndex: 0,
			edit: { tickLabelPosition: 'none' },
		});
		const patched = await savedPart(book);
		expect(patched).toContain('uri="retained"');
		expect(
			parseChart(patched, model.anchor, '').formatting?.entries.categoryAxis?.labelsVisible,
		).toBe(false);
	});

	it.each(['type', 'grouping', 'series'] as const)(
		'retains flags when a %s change edits a kept chart',
		async (change) => {
			const model = chart('column', true, false);
			const book = await withChartPart(chartXml(model));
			const live = book.sheets[0]!.drawings[0] as ChartObject;
			const patch: ChartPatch =
				change === 'type'
					? { chartType: 'line' }
					: change === 'grouping'
						? { grouping: 'stacked' }
						: { series: [...live.series, { categories: ['A', 'B'], values: [30, 40] }] };
			createEditSession(book).updateChart(0, 0, patch);
			const back = parseChart(await savedPart(book), model.anchor, '');
			expect(back.formatting?.entries.categoryAxis?.labelsVisible).toBe(false);
			expect(back.formatting?.entries.valueAxis?.labelsVisible).toBe(false);
		},
	);
});
