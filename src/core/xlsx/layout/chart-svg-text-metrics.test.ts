import { expect, it, vi } from 'vitest';
import native from '../../chart/excel-chart-title-text.json';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { chartTextWidth, fitChartText } from './chart-svg-text-metrics';
import { parseXml } from '../../xml';

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
