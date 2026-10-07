import { expect, it } from 'vitest';
import { parseXml } from '../../xml';
import { createWorkbook } from '../workbook';
import type { ChartObject } from '../model';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { chartAreaRect } from './chart-svg-appearance';

function chart(): ChartObject {
	return {
		kind: 'chart',
		chartType: 'column',
		title: 'Sales chart',
		showLegend: true,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [{ name: 'Sales', categories: ['A', 'B'], values: [10, 20] }],
	};
}

it('omits backgrounds with no paintable layout area', () => {
	const view = chartView(createWorkbook(), 0, chart(), () => []);
	view.appearance = { legend: { gradient: { type: 'radial', path: 'circle', stops: [] } } };
	for (const area of [
		{ x: 0, y: 0, w: 0, h: 20 },
		{ x: 0, y: 0, w: 20, h: -4 },
	]) {
		expect(chartAreaRect(view, 'legend', area)).toBe('');
	}
});

it.each(['l', 'r', 't', 'b', 'tr'] as const)(
	'paints title and %s legend behind text without changing the view',
	(position) => {
		const source = chart();
		source.legendPosition = position;
		source.formatting = {
			sourceXml: '',
			entries: {
				title: {
					sourceXml: '',
					fill: {
						kind: 'solid',
						color: {
							kind: 'srgb',
							value: 'FF0000',
							transforms: [{ name: 'alpha', value: '63000' }],
						},
					},
				},
				legend: {
					sourceXml: '',
					fill: { kind: 'solid', color: { kind: 'srgb', value: '00FF00', transforms: [] } },
				},
			},
		};
		const view = chartView(createWorkbook(), 0, source, () => []);
		const before = structuredClone(view);
		const svg = renderChartSvg(view, 600, 400);
		const doc = parseXml(svg);
		for (const [part, label, paint] of [
			['title', 'Sales chart', 'rgba(255,0,0,0.63)'],
			['legend', 'Sales', '#00FF00'],
		] as const) {
			const rect = [...doc.getElementsByTagName('rect')].find(
				(el) => el.getAttribute('data-chart-part') === part,
			)!;
			expect(rect).toBeDefined();
			expect(rect.getAttribute('fill')).toBe(paint);
			expect(Number(rect.getAttribute('width'))).toBeGreaterThan(0);
			expect(Number(rect.getAttribute('height'))).toBeGreaterThan(0);
			expect(svg.indexOf(`data-chart-part="${part}"`)).toBeLessThan(
				svg.indexOf(`>${label}</text>`),
			);
		}
		expect(view).toEqual(before);
	},
);

it.each(['rect', 'circle', 'shape', undefined] as const)(
	'uses each decoration rectangle for %s gradient geometry',
	(path) => {
		const source = chart();
		source.formatting = { sourceXml: '', entries: {} };
		for (const part of ['title', 'legend'] as const)
			source.formatting.entries[part] = {
				sourceXml: '',
				fill: {
					kind: 'gradient',
					...(path ? { path } : { angle: 90, scaled: true }),
					stops: [
						{ position: 0, color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
						{ position: 100, color: { kind: 'srgb', value: 'FFFFFF', transforms: [] } },
					],
				},
			};
		const doc = parseXml(
			renderChartSvg(
				chartView(createWorkbook(), 0, source, () => []),
				600,
				400,
			),
		);
		for (const part of ['title', 'legend'] as const) {
			const rect = [...doc.getElementsByTagName('rect')].find(
				(el) => el.getAttribute('data-chart-part') === part,
			)!;
			const id = rect.getAttribute('fill')!.slice(5, -1);
			expect(id.endsWith(`-${part}`)).toBe(true);
			const definition = [...doc.getElementsByTagName('*')].find(
				(el) => el.getAttribute('id') === id,
			)!;
			expect(definition).toBeDefined();
			if (path === 'circle') {
				const matrix = definition
					.getAttribute('gradientTransform')!
					.slice(7, -1)
					.split(' ')
					.map(Number);
				expect(matrix[3]).toBeCloseTo(
					Number(rect.getAttribute('width')) / Number(rect.getAttribute('height')),
					2,
				);
				expect(matrix[5]).toBeCloseTo((1 - matrix[3]!) / 2, 5);
			}
		}
		delete source.title;
		source.showLegend = false;
		const hidden = renderChartSvg(
			chartView(createWorkbook(), 0, source, () => []),
			600,
			400,
		);
		expect(hidden).not.toContain('data-chart-part="title"');
		expect(hidden).not.toContain('data-chart-part="legend"');
	},
);
