import { expect, it } from 'vitest';
import { parseXml } from '../../xml';
import type { ChartViewModel } from './chart-view';
import { renderChartSvg } from '../../chart/render/chart-svg';

for (const horizontal of [false, true])
	for (const grouping of ['clustered', 'stacked', 'percentStacked'] as const)
		it(`keeps circular mark paints independent in ${horizontal ? 'bar' : 'column'} ${grouping}`, () => {
			const gradient = {
				type: 'radial' as const,
				path: 'circle',
				stops: [
					{ position: 0, color: '#ff0000' },
					{ position: 100, color: '#ffffff' },
				],
			};
			const model: ChartViewModel = {
				type: horizontal ? 'bar' : 'column',
				horizontal,
				grouping,
				showLegend: true,
				legendPosition: 'r',
				supported: true,
				categories: ['A', 'B', 'C', 'D'],
				valueAxis: {
					min: -5,
					max: 10,
					majorUnit: 5,
					ticks: [-5, 0, 5, 10],
					percent: grouping === 'percentStacked',
				},
				series: [
					{ name: 'First', values: [1, -2, 0, null], color: '#ff0000', gradient },
					{
						name: 'Second',
						values: [3, 4, 2, 1],
						color: '#ff0000',
						gradient,
						pointColors: ['#00ff00'],
						pointGradients: { 1: { ...gradient, focalPoint: { x: 1, y: 0 } } },
					},
				],
			};
			const before = structuredClone(model),
				doc = parseXml(renderChartSvg(model, 960, 600));
			const groups = [...doc.getElementsByTagName('g')].filter((g) =>
				g.hasAttribute('data-chart-series'),
			);
			expect(groups).toHaveLength(7);
			const ids = new Set<string>();
			for (const group of groups) {
				const rect = group.getElementsByTagName('rect')[0]!;
				const width = Number(rect.getAttribute('width')),
					height = Number(rect.getAttribute('height'));
				if (
					group.getAttribute('data-chart-series') === '1' &&
					group.getAttribute('data-chart-point') === '0'
				) {
					expect(rect.getAttribute('fill')).toBe('#00ff00');
					continue;
				}
				const id = rect.getAttribute('fill')!.slice(5, -1);
				expect(ids.has(id)).toBe(false);
				ids.add(id);
				const def = [...doc.getElementsByTagName('radialGradient')].find(
					(node) => node.getAttribute('id') === id,
				)!;
				expect(def).toBeDefined();
				if (width === 0 || height === 0) continue;
				const matrix = def.getAttribute('gradientTransform')!.slice(7, -1).split(' ').map(Number);
				expect(matrix[3]).toBeCloseTo(width / height, 3);
				if (
					group.getAttribute('data-chart-series') === '1' &&
					group.getAttribute('data-chart-point') === '1'
				) {
					expect(id).toMatch(/-s1-p1$/);
					expect(def.getAttribute('cx')).toBe('1');
					expect(def.getAttribute('cy')).toBe('0');
				} else expect(id).toMatch(/-s[01]$/);
			}
			expect(model).toEqual(before);
		});
