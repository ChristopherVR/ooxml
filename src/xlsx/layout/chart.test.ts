import { describe, expect, it } from 'vitest';
import type { CellValue, ChartObject, ChartType } from '../model.js';
import { createWorkbook, DEFAULT_THEME } from '../workbook.js';
import { autoSeriesColor, modulateLuminance } from './chart-colors.js';
import { formatAxisValue, niceScale } from './chart-scale.js';
import { renderChartSvg } from './chart-svg.js';
import { esc } from './chart-svg-util.js';
import { chartView, type ChartViewModel } from './chart-view.js';

const anchor = { from: { row: 0, col: 0, colOffset: 0, rowOffset: 0 } };

function chart(type: ChartType, extra: Partial<ChartObject> = {}): ChartObject {
	return {
		kind: 'chart',
		anchor,
		chartType: type,
		showLegend: true,
		series: [
			{ name: 'North', categories: ['Q1', 'Q2', 'Q3'], values: [10, 20, 30] },
			{ name: 'South', categories: ['Q1', 'Q2', 'Q3'], values: [5, 15, 25] },
		],
		...extra,
	};
}

const noRefs = (): CellValue[] => [];

describe('niceScale', () => {
	it.each([
		[0, 100, 0, 120, 20],
		[1, 10, 0, 12, 2],
		[0, 50, 0, 60, 10],
		[0, 7, 0, 8, 1],
		[0, 1, 0, 1.2, 0.2],
		[-30, 70, -40, 80, 20],
		[-50, -10, -60, 0, 10],
		[95, 100, 94, 101, 1],
		[0, 0, 0, 1.2, 0.2],
		[0, 1234567, 0, 1400000, 200000],
	])('scales %s..%s to %s..%s step %s', (lo, hi, min, max, unit) => {
		const s = niceScale(lo, hi);
		expect(s.min).toBe(min);
		expect(s.max).toBe(max);
		expect(s.majorUnit).toBe(unit);
		expect(s.ticks[0]).toBe(min);
		expect(s.ticks.at(-1)).toBe(max);
		expect(s.ticks.length - 1).toBeLessThanOrEqual(10);
	});

	it('swaps reversed bounds and handles a single negative value', () => {
		expect(niceScale(10, 0).max).toBe(12);
		const neg = niceScale(-5, -5);
		expect(neg.max).toBe(0);
		expect(neg.min).toBeLessThan(-5);
	});

	it('honours maxIntervals and noZero', () => {
		expect(niceScale(0, 100, { maxIntervals: 5 }).majorUnit).toBe(50);
		expect(niceScale(50, 100, { noZero: true }).min).toBeGreaterThan(0);
	});

	it('formats axis values', () => {
		expect(formatAxisValue(1200000)).toBe('1,200,000');
		expect(formatAxisValue(0.30000000000000004)).toBe('0.3');
		expect(formatAxisValue(-2500.5)).toBe('-2,500.5');
		expect(formatAxisValue(0.4, true)).toBe('40%');
		expect(formatAxisValue(1e20)).toBe('1.00E+20');
	});
});

describe('series colours', () => {
	it('uses theme accents 1-6 in order', () => {
		expect([0, 1, 2, 3, 4, 5].map((i) => autoSeriesColor(DEFAULT_THEME, i))).toEqual([
			'#4472C4',
			'#ED7D31',
			'#A5A5A5',
			'#FFC000',
			'#5B9BD5',
			'#70AD47',
		]);
	});

	it('darkens then lightens the accents for later rounds', () => {
		const seventh = autoSeriesColor(DEFAULT_THEME, 6);
		expect(seventh).toBe(modulateLuminance('4472C4', 0.6));
		expect(seventh).not.toBe('#4472C4');
		const thirteenth = autoSeriesColor(DEFAULT_THEME, 12);
		expect(thirteenth).toBe(modulateLuminance('4472C4', 0.8, 0.2));
		expect(
			new Set(Array.from({ length: 24 }, (_, i) => autoSeriesColor(DEFAULT_THEME, i))).size,
		).toBe(24);
	});

	it('modulates luminance', () => {
		expect(modulateLuminance('FFFFFF', 0.5)).toBe('#808080');
		expect(modulateLuminance('000000', 1, 0.5)).toBe('#808080');
		expect(modulateLuminance('bad', 1)).toBe('#000000');
	});
});

describe('chartView', () => {
	const wb = createWorkbook();

	it('falls back to cached values and theme colours', () => {
		const m = chartView(wb, 0, chart('column'), noRefs);
		expect(m.categories).toEqual(['Q1', 'Q2', 'Q3']);
		expect(m.series.map((s) => [s.name, s.color])).toEqual([
			['North', '#4472C4'],
			['South', '#ED7D31'],
		]);
		expect(m.grouping).toBe('clustered');
		expect(m.valueAxis).toMatchObject({ min: 0, max: 35, majorUnit: 5, percent: false });
		expect(m.horizontal).toBe(false);
		expect(m.supported).toBe(true);
	});

	it('resolves series from live references', () => {
		const refs: Record<string, CellValue[]> = {
			'Sheet1!$B$1': ['Live'],
			'Sheet1!$A$2:$A$4': ['Jan', 'Feb', 'Mar'],
			'Sheet1!$B$2:$B$4': [100, 'n/a', 300],
		};
		const c = chart('line', {
			series: [
				{
					nameRef: 'Sheet1!$B$1',
					categoriesRef: 'Sheet1!$A$2:$A$4',
					valuesRef: 'Sheet1!$B$2:$B$4',
					categories: [],
					values: [1, 2, 3],
				},
			],
		});
		const m = chartView(wb, 0, c, (ref) => refs[ref] ?? []);
		expect(m.series[0]).toMatchObject({ name: 'Live', values: [100, null, 300] });
		expect(m.categories).toEqual(['Jan', 'Feb', 'Mar']);
		expect(m.grouping).toBe('standard');
	});

	it('survives a throwing reference resolver', () => {
		const m = chartView(wb, 0, chart('bar'), () => {
			throw new Error('bad ref');
		});
		expect(m.series[0]?.values).toEqual([10, 20, 30]);
		expect(m.horizontal).toBe(true);
	});

	it('scales stacked and percent-stacked charts', () => {
		expect(
			chartView(wb, 0, chart('column', { grouping: 'stacked' }), noRefs).valueAxis,
		).toMatchObject({ min: 0, max: 60 });
		const pct = chartView(wb, 0, chart('area', { grouping: 'percentStacked' }), noRefs).valueAxis;
		expect(pct).toMatchObject({ min: 0, max: 1, percent: true });
	});

	it('gives pie points their own colours and no axis', () => {
		const m = chartView(wb, 0, chart('pie'), noRefs);
		expect(m.series[0]?.pointColors).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
		expect(m.valueAxis).toBeUndefined();
	});

	it('uses explicit series colours and default names', () => {
		const m = chartView(
			wb,
			0,
			chart('line', { series: [{ categories: [], values: [1, 2], color: { rgb: 'FF0000' } }] }),
			noRefs,
		);
		expect(m.series[0]).toMatchObject({ name: 'Series1', color: '#FF0000' });
		expect(m.categories).toEqual(['1', '2']);
	});

	it('builds numeric X values for scatter charts', () => {
		const m = chartView(
			wb,
			0,
			chart('scatter', { series: [{ categories: [1, 2, 4], values: [3, 5, 9] }] }),
			noRefs,
		);
		expect(m.series[0]?.xValues).toEqual([1, 2, 4]);
		expect(m.xAxis).toMatchObject({ min: 0 });
		const text = chartView(
			wb,
			0,
			chart('scatter', { series: [{ categories: ['a', 'b'], values: [3, 5] }] }),
			noRefs,
		);
		expect(text.series[0]?.xValues).toEqual([1, 2]);
	});

	it('flags unsupported types honestly', () => {
		expect(chartView(wb, 0, chart('bubble'), noRefs).supported).toBe(false);
	});
});

describe('renderChartSvg', () => {
	const wb = createWorkbook();
	const render = (c: ChartObject): string => renderChartSvg(chartView(wb, 0, c, noRefs), 400, 300);
	const count = (svg: string, tag: string): number => svg.split(`<${tag} `).length - 1;

	it('produces a standalone SVG document', () => {
		const svg = render(chart('column', { title: 'Sales' }));
		expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300"')).toBe(
			true,
		);
		expect(svg.endsWith('</svg>')).toBe(true);
		expect(svg).toContain('>Sales</text>');
		expect(svg).toContain('<title>Sales</title>');
		expect(svg).not.toMatch(/NaN|undefined/);
	});

	it('draws one bar per point plus legend swatches', () => {
		const svg = render(chart('column'));
		// background + 6 bars + 2 legend swatches
		expect(count(svg, 'rect')).toBe(9);
		expect(svg).toContain('>North</text>');
		expect(svg).toContain('>Q2</text>');
		expect(svg).toContain('>35</text>');
	});

	it('draws stacked and horizontal bars', () => {
		expect(count(render(chart('bar', { grouping: 'stacked' })), 'rect')).toBe(9);
		expect(count(render(chart('column', { grouping: 'percentStacked' })), 'rect')).toBe(9);
		expect(render(chart('column', { grouping: 'percentStacked' }))).toContain('>100%</text>');
	});

	it('draws lines with markers and areas as polygons', () => {
		const line = render(chart('line'));
		expect(count(line, 'polyline')).toBe(2);
		expect(count(line, 'circle')).toBe(6);
		expect(count(render(chart('area')), 'polygon')).toBe(2);
	});

	it('breaks lines at gaps', () => {
		const svg = render(
			chart('line', {
				series: [{ name: 'A', categories: ['a', 'b', 'c', 'd'], values: [1, 2, null, 4] }],
			}),
		);
		expect(count(svg, 'polyline')).toBe(1);
		expect(count(svg, 'circle')).toBe(3);
	});

	it('draws pie slices, doughnut rings and radar polygons', () => {
		expect(count(render(chart('pie')), 'path')).toBe(3);
		expect(count(render(chart('doughnut')), 'path')).toBe(6);
		const radar = render(chart('radar'));
		expect(count(radar, 'polygon')).toBeGreaterThan(2);
		const single = render(
			chart('pie', { series: [{ name: 'A', categories: ['x'], values: [5] }] }),
		);
		expect(count(single, 'path')).toBe(1);
	});

	it('draws scatter markers', () => {
		const svg = render(
			chart('scatter', { series: [{ name: 'S', categories: [1, 2, 3], values: [2, 4, 8] }] }),
		);
		expect(count(svg, 'circle')).toBe(3);
	});

	it('escapes every piece of text', () => {
		const svg = render(
			chart('column', {
				title: '<b>&"\'</b>',
				series: [{ name: '<script>', categories: ['a&b'], values: [1] }],
			}),
		);
		expect(svg).not.toContain('<script>');
		expect(svg).not.toContain('<b>');
		expect(svg).toContain('&lt;b&gt;&amp;&quot;&#39;&lt;/b&gt;');
		expect(svg).toContain('a&amp;b');
		expect(esc('<>&"\'')).toBe('&lt;&gt;&amp;&quot;&#39;');
	});

	it('places the legend at the bottom or hides it', () => {
		const bottom = render(chart('line', { legendPosition: 'b' }));
		expect(bottom).toContain('>South</text>');
		const hidden = render(chart('line', { showLegend: false }));
		expect(hidden).not.toContain('>South</text>');
	});

	it('labels unsupported charts instead of drawing them', () => {
		expect(render(chart('surface'))).toContain('surface charts are not drawn');
	});

	it('handles empty models and tiny sizes', () => {
		const empty: ChartViewModel = chartView(wb, 0, chart('column', { series: [] }), noRefs);
		expect(renderChartSvg(empty, 10, 10)).toContain('</svg>');
		expect(renderChartSvg(chartView(wb, 0, chart('pie'), noRefs), 0, 0)).toContain('width="1"');
	});
});
