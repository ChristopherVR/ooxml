import { expect, it } from 'vitest';
import { parseXml } from '../../xml';
import type { ChartViewModel } from './chart-view';
import { renderChartSvg } from './chart-svg';

function view(path: string): ChartViewModel {
	const gradient = {
		type: 'radial' as const,
		path,
		stops: [
			{ position: 0, color: '#ff0000' },
			{ position: 100, color: '#ffffff' },
		],
	};
	return {
		type: 'column',
		grouping: 'clustered',
		horizontal: false,
		supported: true,
		showLegend: true,
		legendPosition: 'r',
		title: 'Bounds',
		categories: ['A'],
		series: [{ name: 'Series', values: [1], color: '#00ff00' }],
		valueAxis: { min: 0, max: 2, ticks: [0, 1, 2], majorUnit: 1, percent: false },
		appearance: { chartArea: { gradient }, plotArea: { gradient } },
	};
}

it('sizes circular chart and plot paints independently after title, legend and axes layout', () => {
	const model = view('circle'),
		original = structuredClone(model);
	const svg = parseXml(renderChartSvg(model, 960, 600));
	const gradients = [...svg.getElementsByTagName('radialGradient')];
	expect(gradients).toHaveLength(2);
	for (const gradient of gradients) {
		const rectangle = [...svg.getElementsByTagName('rect')].find(
			(rect) => rect.getAttribute('fill') === `url(#${gradient.getAttribute('id')})`,
		)!;
		const width = Number(rectangle.getAttribute('width')),
			height = Number(rectangle.getAttribute('height'));
		expect(Number(gradient.getAttribute('r'))).toBeCloseTo(
			Math.hypot(0.5, (0.5 * height) / width),
			4,
		);
		const matrix = gradient.getAttribute('gradientTransform')!.slice(7, -1).split(' ').map(Number);
		expect(matrix[3]).toBeCloseTo(width / height, 4);
		expect(matrix[5]).toBeCloseTo(0.5 * (1 - width / height), 4);
	}
	expect(gradients[0]!.getAttribute('gradientTransform')).not.toBe(
		gradients[1]!.getAttribute('gradientTransform'),
	);
	expect(model).toEqual(original);
});

it('uses the shared rectangular path painter for chart and plot rectangle shape outlines', () => {
	const svg = parseXml(renderChartSvg(view('shape'), 960, 600));
	expect(svg.getElementsByTagName('radialGradient')).toHaveLength(0);
	expect(svg.getElementsByTagName('pattern')).toHaveLength(2);
	for (const pattern of svg.getElementsByTagName('pattern')) {
		const href = pattern.getElementsByTagName('image')[0]!.getAttribute('href')!;
		expect(decodeURIComponent(href)).toContain('<rect');
	}
});
