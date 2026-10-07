// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createChartGradientPreview } from './chart-gradient-preview';

it('paints crossed stops in all frozen panes and restores only the affected series', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const markup = `<div class="xg-obj" data-index="2"><svg><defs><linearGradient id="chart-s1"><stop offset="0.2" stop-color="#FF0000"/><stop offset="0.8" stop-color="#0000FF"/></linearGradient><linearGradient id="chart-s0"><stop offset="0" stop-color="#FFFFFF"/></linearGradient></defs></svg></div>`;
	root.innerHTML = markup + markup.replace('chart-', 'second-');
	const before = root.innerHTML;
	const gradient = {
		type: 'linear' as const,
		stops: [
			{ position: 20, color: '#FF0000', opacity: 0.63 },
			{ position: 80, color: '#0000FF' },
		],
	};
	const preview = createChartGradientPreview(root, 2, 1, gradient);
	preview.position(0, 85);
	expect(gradient.stops[0]!.position).toBe(20);
	for (const node of root.querySelectorAll('linearGradient[id$="-s1"]')) {
		expect(node.children[0]!.getAttribute('offset')).toBe('0.8');
		expect(node.children[1]!.getAttribute('stop-color')).toBe('#FF0000');
		expect(node.children[1]!.getAttribute('stop-opacity')).toBe('0.63');
	}
	preview.restore();
	expect(root.innerHTML).toBe(before);
	host.remove();
});

it('previews a background in every frozen pane while preserving physical radial bounds and other paints', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const markup = `<div class="xg-obj" data-index="2"><svg><defs><radialGradient id="chart-chartArea" cx="1" cy="1" r="1.2" gradientTransform="matrix(1 0 0 2 0 -1)"><stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#FFFFFF"/></radialGradient><linearGradient id="chart-plotArea"><stop offset="0" stop-color="#0000FF"/></linearGradient></defs></svg></div>`;
	root.innerHTML = markup + markup.replace(/chart-/g, 'second-');
	const before = root.innerHTML;
	const preview = createChartGradientPreview(root, 2, 'chartArea', {
		type: 'radial',
		path: 'circle',
		focalPoint: { x: 1, y: 1 },
		stops: [
			{ position: 0, color: '#FF0000' },
			{ position: 100, color: '#FFFFFF' },
		],
	});
	preview.position(0, 40);
	for (const radial of root.querySelectorAll('radialGradient')) {
		expect(radial.getAttribute('gradientTransform')).toBe('matrix(1 0 0 2 0 -1)');
		expect(radial.getAttribute('r')).toBe('1.2');
		expect(radial.children[0]!.getAttribute('offset')).toBe('0.4');
	}
	for (const plot of root.querySelectorAll('linearGradient'))
		expect(plot.children[0]!.getAttribute('stop-color')).toBe('#0000FF');
	preview.restore();
	expect(root.innerHTML).toBe(before);
	host.remove();
});
