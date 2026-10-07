import { expect, it } from 'vitest';
import native from '../../chart/excel-chart-styles.json';
import { readChartStyle } from '../../chart/read-style';
import { parseXml } from '../../xml/index';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { THEME_SLOTS } from './colors';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { chartPointsToPixels } from './chart-appearance';

for (const sample of native.cases)
	it(`renders native style ${sample.style} text measured by Excel COM`, () => {
		const book = createWorkbook();
		book.theme.colors = THEME_SLOTS.map((slot) =>
			(native.scheme as Record<string, string>)[slot]!.slice(1),
		);
		book.theme.majorFont = 'Aptos Display';
		book.theme.minorFont = 'Aptos Narrow';
		const chart = parseChart(
			sample.parts['xl/charts/chart1.xml'],
			{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
			'xl/charts/chart1.xml',
		);
		chart.styleDefinition = readChartStyle(sample.parts['xl/charts/style1.xml'])!;
		const view = chartView(book, 0, chart, () => []);
		const svg = parseXml(renderChartSvg(view, 640, 400));
		const texts = [...svg.getElementsByTagName('text')];
		for (const [part, content, font] of [
			['title', 'Native style', sample.titleText],
			['legend', 'Sales', sample.legendText],
			['categoryAxis', 'M2', sample.categoryText],
			['valueAxis', '0', sample.valueText],
		] as const) {
			const text = texts.find((text) => text.textContent === content)!;
			if (part === 'valueAxis' && !sample.hasValueAxis) {
				expect(text, 'deleted native value axis').toBeUndefined();
				continue;
			}
			expect(text, part).toBeDefined();
			expect(text.getAttribute('font-family')?.split(',')[0]?.replaceAll('"', ''), part).toBe(
				font.name,
			);
			if (font.size !== null)
				expect(Number(text.getAttribute('font-size')), part).toBeCloseTo(
					chartPointsToPixels(font.size),
					8,
				);
			expect(text.getAttribute('fill')?.toUpperCase(), part).toBe(font.color);
			expect(text.getAttribute('font-weight') === 'bold', part).toBe(font.bold);
			expect(text.getAttribute('font-style') === 'italic', part).toBe(font.italic);
		}
		const area = view.appearance?.chartArea;
		expect((area?.gradient?.stops[0]?.color ?? area?.fillColor)?.toUpperCase()).toBe(
			sample.chartArea.fillColor,
		);
		if (area?.gradient) {
			expect(svg.getElementsByTagName('defs')).toHaveLength(1);
			expect(svg.getElementsByTagName('stop')).toHaveLength(area.gradient.stops.length);
		}
		if (sample.chartArea.lineWeight !== null) {
			expect(view.appearance?.chartArea?.lineWidth).toBe(sample.chartArea.lineWeight);
			expect(view.appearance?.chartArea?.lineColor?.toUpperCase()).toBe(sample.chartArea.lineColor);
		}
		if (sample.style === 209)
			expect(view.appearance?.gridlineMajor?.lineColor).toBe('rgba(242,242,242,0.1)');
	});
