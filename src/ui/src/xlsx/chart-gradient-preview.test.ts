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
