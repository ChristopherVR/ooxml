import { describe, it, expect } from 'vitest';
import { demoDocument } from './demo-document.js';
import { renderPage } from './render-svg.js';
describe('normalized linear gradients', () => {
	it('draws local y-up endpoints and per-stop opacity without external paint URLs', () => {
		const model = structuredClone(demoDocument);
		model.pages[0]!.shapes[0]!.style.fillGradient = {
			type: 'linear',
			start: [0, 0],
			end: [2, 1],
			stops: [
				{ offset: 0, color: '#fff', opacity: 1 },
				{ offset: 1, color: '#126a64', opacity: 0.5 },
			],
		};
		const result = renderPage(model, model.pages[0]!);
		const gradient = result.svg.querySelector('linearGradient')!;
		expect(gradient.getAttribute('gradientUnits')).toBe('userSpaceOnUse');
		expect(gradient.getAttribute('y2')).toBe('1');
		expect(gradient.querySelectorAll('stop')).toHaveLength(2);
		expect(result.svg.querySelector('[data-shape-id="s1"] path')?.getAttribute('fill')).toMatch(
			/^url\(#visio-fill-/,
		);
	});
});
