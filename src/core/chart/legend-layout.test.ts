import { expect, it } from 'vitest';
import native from './excel-chart-legend-layout.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { chartLegendSvg } from './render/chart-svg-legend';
import { renderChartSvg } from './render/chart-svg';
import { chartXml } from '../xlsx/write/chart';
import { NS, first, parseXml, buildXml } from '../xml';

// Chromium 154 canvas, Arial 16 CSS px, captured through Playwright MCP on Windows.
const widths: Record<string, number> = { Sales: 40.0234375, Costs: 40.8984375, Profit: 37.34375 };
const options = {
	measureText: (value: string) => widths[value] ?? value.length * 8,
	measureFont: () => ({ ascent: 14, descent: 3 }),
};
const legendProperty = (xml: string, name: string) => {
	const legend = first(first(parseXml(xml).documentElement, 'chart', NS.c), 'legend', NS.c);
	const node = first(legend, name, NS.c);
	return node ? buildXml(node) : undefined;
};

for (const sample of native.cases)
	it(`renders and preserves native legend layout ${sample.referenceName}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const book = createWorkbook();
		const frame = {
			width: (sample.layoutGeometry.chart.widthPt * 4) / 3,
			height: (sample.layoutGeometry.chart.heightPt * 4) / 3,
		};
		const area = { x: 8, y: 40, w: frame.width - 16, h: frame.height - 48 };
		const verify = (target: typeof chart) => {
			const model = chartView(book, 0, target, () => []);
			expect(model.legendOverlay).toBe(!sample.layoutGeometry.legendIncludeInLayout);
			const painted = chartLegendSvg(model, area, frame, options);
			const { legendLayout: _layout, ...automaticModel } = model;
			const automatic = chartLegendSvg(
				{ ...automaticModel, legendOverlay: false },
				area,
				frame,
				options,
			);
			expect(painted.plot).toEqual(model.legendOverlay ? area : automatic.plot);
			if (sample.referenceName.endsWith('auto')) {
				expect(model.legendLayout).toBeUndefined();
				return;
			}
			const svg = parseXml(renderChartSvg(model, frame.width, frame.height, options));
			const group = [...svg.getElementsByTagName('g')].find(
				(node) => node.getAttribute('data-chart-legend-layout') === 'true',
			)!;
			expect(group).toBeDefined();
			const box = group.getElementsByTagName('rect')[0]!;
			const nativeBox = sample.layoutGeometry.legend;
			// The source manual x/y fractions use COM legend coordinates +4pt in this corpus.
			for (const [attribute, value] of [
				['x', nativeBox.leftPt + 4],
				['y', nativeBox.topPt + 4],
				['width', nativeBox.widthPt],
				['height', nativeBox.heightPt],
			] as const)
				expect(Number(box.getAttribute(attribute))).toBeCloseTo((value * 4) / 3, 1);
			const labels = [...group.getElementsByTagName('text')];
			expect(labels.map((node) => node.textContent)).toEqual(['Sales', 'Costs', 'Profit']);
			const xs = labels.map((node) => Number(node.getAttribute('x')));
			const ys = labels.map((node) => Number(node.getAttribute('y')));
			if (sample.referenceName.includes('bottom') || sample.referenceName.includes('wide')) {
				expect(new Set(ys).size).toBe(1);
				expect(xs[0]).toBeLessThan(xs[1]!);
				expect(xs[1]).toBeLessThan(xs[2]!);
				// Approximate native PNG text origins converted from 144dpi to CSS px.
				// These bounded checks expose spacing errors without claiming exact raster parity.
				const nativeXs = sample.referenceName.includes('wide')
					? [115.33, 210, 305.33]
					: [84, 149.33, 214.67];
				for (let i = 0; i < xs.length; i++) expect(Math.abs(xs[i]! - nativeXs[i]!)).toBeLessThan(6);
			} else {
				expect(new Set(xs).size).toBe(1);
				expect(ys[1]! - ys[0]!).toBeCloseTo((nativeBox.heightPt * 4) / 9, 1);
				expect(ys[2]! - ys[1]!).toBeCloseTo((nativeBox.heightPt * 4) / 9, 1);
			}
			expect(group.nextSibling).toBeNull();
		};
		verify(chart);
		const saved = chartXml({ ...chart, chartType: 'line' });
		for (const name of ['layout', 'overlay'])
			expect(legendProperty(saved, name)).toBe(legendProperty(source, name));
		verify(parseChart(saved, chart.anchor, ''));
	});
