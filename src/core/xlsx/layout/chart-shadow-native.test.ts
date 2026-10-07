import { expect, it } from 'vitest';
import native from '../__fixtures__/excel-chart-gradients.json';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { THEME_SLOTS } from './colors';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { NS, parseXml } from '../../xml/index';

for (const sample of native.cases)
	it(`renders palette ${sample.palette} shadows measured by Excel COM`, () => {
		const book = createWorkbook();
		book.theme.colors = THEME_SLOTS.map((slot) =>
			(native.scheme as Record<string, string>)[slot]!.slice(1),
		);
		const chart = parseChart(
			sample.parts['xl/charts/chart1.xml'],
			{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			'xl/charts/chart1.xml',
		);
		const view = chartView(book, 0, chart, () => []);
		for (const [i, series] of view.series.entries()) {
			const measured = sample.series[i]!.shadow;
			expect(measured.visible).toBe(true);
			expect(series.shadow?.color).toBe(measured.color);
			expect(series.shadow?.blur).toBeCloseTo((measured.blurPoints * 4) / 3, 6);
			expect(series.shadow?.dx).toBeCloseTo((measured.offsetXPoints * 4) / 3, 6);
			expect(series.shadow?.dy).toBeCloseTo((measured.offsetYPoints * 4) / 3, 6);
			expect(series.shadow?.opacity).toBeCloseTo(measured.opacity, 6);
		}
		const title = sample.titleShadow;
		expect(view.appearance?.title?.textShadow?.blur).toBeCloseTo((title.blurPoints * 4) / 3, 6);
		expect(view.appearance?.title?.textShadow?.dy).toBeCloseTo((title.offsetYPoints * 4) / 3, 6);
		expect(view.appearance?.title?.textShadow?.opacity).toBeCloseTo(title.opacity, 6);
		const svg = parseXml(renderChartSvg(view, 640, 400));
		expect(svg.getElementsByTagName('feDropShadow')).toHaveLength(3);
		const filters = [...svg.getElementsByTagName('filter')];
		expect(filters.every((filter) => filter.getAttribute('filterUnits') === 'userSpaceOnUse')).toBe(
			true,
		);
		const titleNode = [...svg.getElementsByTagName('text')].find(
			(text) => text.textContent === 'Native style',
		)!;
		expect(titleNode.getAttribute('filter')).toContain('-title-text-shadow)');
		expect(
			[...svg.getElementsByTagName('rect')].filter((rect) =>
				rect.getAttribute('filter')?.includes('-s0-shadow)'),
			),
		).toHaveLength(4);
		expect(view.series[0]!.shadowFilter).toBeUndefined();
		expect(chart.series[0]!.effectsXml).toContain('outerShdw');
		expect(parseXml(chart.series[0]!.effectsXml!).documentElement.namespaceURI).toBe(NS.a);
	});
