import { expect, it, vi } from 'vitest';
import native from '../../chart/excel-chart-title-text.json';
import geometry from '../../chart/excel-chart-title-geometry.json';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { chartView } from './chart-view';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { chartTextWidth, fitChartText } from '../../chart/render/chart-svg-text-metrics';
import { parseXml } from '../../xml';
import { chartRichTitleSvg } from '../../chart/render/chart-svg-title-text';

it('places native mixed runs contiguously and centers their measured line', () => {
	const source = native.cases.find((item) => item.referenceName === 'mixed-lines')!;
	const chart = parseChart(
		source.parts['xl/charts/chart1.xml'],
		{
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
		},
		'',
	);
	const book = createWorkbook();
	book.theme.minorFont = 'Aptos Narrow';
	const widths: Record<string, number> = { Revenue: 133, ' ': 5, growth: 46, Forecast: 97 };
	const measureText = vi.fn((content: string, _font: string) => widths[content] ?? 10);
	const svg = parseXml(
		renderChartSvg(
			chartView(book, 0, chart, () => []),
			640,
			400,
			{ measureText },
		),
	);
	const runs = [...svg.getElementsByTagName('tspan')];
	const x = (index: number) => Number(runs[index]!.getAttribute('x'));
	expect(x(0)).toBe((640 - 184) / 2);
	expect(x(1) - x(0)).toBe(133);
	expect(x(2) - x(1)).toBe(5);
	expect(x(3)).toBe((640 - 97) / 2);
	expect(measureText.mock.calls).toContainEqual([
		'Revenue',
		'normal bold 32px Arial, "Liberation Sans", Helvetica, sans-serif',
	]);
	expect(
		measureText.mock.calls.some(
			([content, font]) => content === 'growth' && font.startsWith('italic normal 16px'),
		),
	).toBe(true);
});

it('fits by measured advances without cutting a Unicode code point', () => {
	const options = {
		measureText: (content: string) =>
			[...content].reduce((sum, ch) => sum + (ch === 'W' ? 20 : 5), 0),
	};
	expect(fitChartText('WW😀ii', 30, { size: 12 }, options)).toBe('W…');
	expect(fitChartText('😀iiWW', 25, { size: 12 }, options)).toBe('😀ii…');
	expect(fitChartText('iiii', 25, { size: 12 }, options)).toBe('iiii');
});

it.each([NaN, Infinity, -1])(
	'retains finite fallback geometry for an invalid host width %s',
	(width) => {
		expect(chartTextWidth('abc', { size: 20 }, { measureText: () => width })).toBeCloseTo(33);
	},
);

// Font boxes observed with the owned native references in Chromium on Windows.
const measureFont = (font: string) =>
	font.includes('32px Arial')
		? { ascent: 29, descent: 7 }
		: font.includes('Cambria')
			? { ascent: 25, descent: 6 }
			: font.includes('24px')
				? { ascent: 21, descent: 5 }
				: { ascent: 14, descent: 3 };

for (const sample of native.cases)
	it(`uses natural font boxes for native title height ${sample.referenceName}`, () => {
		const chart = parseChart(
			sample.parts['xl/charts/chart1.xml'],
			{
				from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
			},
			'',
		);
		const book = createWorkbook();
		book.theme.minorFont = 'Aptos Narrow';
		const view = chartView(book, 0, chart, () => []);
		const title = chartRichTitleSvg(view, 640, { measureFont })!;
		const reference = geometry.cases.find((item) => item.referenceName === sample.referenceName)!;
		const nativeHeight = (reference.titleGeometry.heightPt * 4) / 3;
		expect(Math.abs(title.height - nativeHeight)).toBeLessThan(2);
		const spans = [...parseXml(`<svg>${title.markup}</svg>`).getElementsByTagName('tspan')];
		expect(Number(spans[0]!.getAttribute('y'))).toBe(40);
		if (sample.referenceName !== 'mixed-single') {
			const forecast = spans.find((span) => span.textContent === 'Forecast')!;
			expect(Number(forecast.getAttribute('y'))).toBe(72);
			const estimated = chartRichTitleSvg(view, 640)!;
			expect(Math.abs(title.height - nativeHeight)).toBeLessThan(
				Math.abs(estimated.height - nativeHeight),
			);
		}
		const fallback = chartRichTitleSvg(view, 640, {
			measureFont: () => ({ ascent: NaN, descent: 7 }),
		});
		expect(fallback).toEqual(chartRichTitleSvg(view, 640));
	});

it('uses the same natural font box for a plain title without changing its baseline', () => {
	const source = native.cases[0]!;
	const chart = parseChart(
		source.parts['xl/charts/chart1.xml'],
		{
			from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
		},
		'',
	);
	const view = chartView(createWorkbook(), 0, { ...chart, title: 'Revenue' }, () => []);
	expect(view.titleText).toBeUndefined();
	view.appearance = {
		...view.appearance,
		title: { ...view.appearance?.title, fillColor: '#FFFFFF' },
	};
	const svg = parseXml(renderChartSvg(view, 640, 400, { measureFont }));
	const box = [...svg.getElementsByTagName('rect')].find(
		(node) => node.getAttribute('data-chart-part') === 'title',
	)!;
	expect(Number(box.getAttribute('height'))).toBe(40);
	const label = [...svg.getElementsByTagName('text')].find(
		(node) => node.textContent === 'Revenue',
	)!;
	expect(Number(label.getAttribute('y'))).toBe(40);
});
