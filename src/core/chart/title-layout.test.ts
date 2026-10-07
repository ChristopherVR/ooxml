import { expect, it } from 'vitest';
import native from './excel-chart-title-layout.json';
import browser from './excel-chart-title-wrap-browser-metrics.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { renderChartSvg } from '../xlsx/layout/chart-svg';
import { chartXml } from '../xlsx/write/chart';
import { NS, first, parseXml, buildXml } from '../xml';

const fonts = browser.fonts as Record<
	string,
	{ ascent: number; descent: number; widths: Record<string, number> }
>;
const regionProperty = (xml: string, region: string, name: string) => {
	const node = first(
		first(first(parseXml(xml).documentElement, 'chart', NS.c), region, NS.c),
		name,
		NS.c,
	);
	return node ? buildXml(node) : undefined;
};

for (const sample of native.cases)
	it(`renders and preserves native title placement ${sample.referenceName}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const book = createWorkbook();
		const verify = (target: typeof chart) => {
			const model = chartView(book, 0, target, () => []);
			expect(model.titleOverlay).toBe(!sample.layoutGeometry.titleIncludeInLayout);
			const width = (sample.layoutGeometry.chart.widthPt * 4) / 3;
			const height = (sample.layoutGeometry.chart.heightPt * 4) / 3;
			const svg = parseXml(
				renderChartSvg(model, width, height, {
					measureText: (text, font) => fonts[font.split(',')[0]!]?.widths[text] ?? 20,
					measureFont: (font) => fonts[font.split(',')[0]!],
				}),
			);
			const group = [...svg.getElementsByTagName('g')].find(
				(node) => node.getAttribute('data-chart-title-layout') === 'true',
			);
			const horizontalLines = [...svg.getElementsByTagName('line')].filter(
				(node) => node.getAttribute('y1') === node.getAttribute('y2'),
			);
			const plotTop = Math.min(...horizontalLines.map((node) => Number(node.getAttribute('y1'))));
			expect(plotTop).toBe(
				model.titleOverlay ? 12 : sample.referenceName === 'layout-moved-long' ? 70 : 43,
			);
			if (sample.referenceName.startsWith('layout-automatic')) {
				expect(group).toBeUndefined();
				return;
			}
			expect(group).toBeDefined();
			const spans = [...group!.getElementsByTagName('tspan')];
			const transform = group!.getAttribute('transform')!.match(/translate\(([^ ]+) ([^)]+)\)/)!;
			const dx = Number(transform[1]),
				dy = Number(transform[2]);
			// For these Excel fixtures, drawing-canvas coordinates are COM title coordinates +4 points.
			const boxX = ((sample.titleGeometry.leftPt + 4) * 4) / 3;
			const boxY = ((sample.titleGeometry.topPt + 4) * 4) / 3;
			const firstX = Number(spans[0]!.getAttribute('x')) + dx;
			const firstY = Number(spans[0]!.getAttribute('y')) + dy;
			expect(firstX).toBeCloseTo(boxX + 4, 1);
			expect(firstY).toBeCloseTo(boxY + 28, 1);
			expect(group!.nextSibling).toBeNull();
		};
		verify(chart);
		const regenerated = chartXml({ ...chart, chartType: 'line' });
		for (const region of ['title', 'plotArea', 'legend'])
			for (const property of ['layout', 'overlay'])
				expect(regionProperty(regenerated, region, property)).toBe(
					regionProperty(source, region, property),
				);
		verify(parseChart(regenerated, chart.anchor, ''));
	});
