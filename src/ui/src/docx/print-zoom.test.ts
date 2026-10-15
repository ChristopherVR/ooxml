import { describe, expect, it } from 'vitest';
import css from './styles/print.css?raw';

describe('Print Layout zoom', () => {
	it('zooms the controller and its nested pages container only once', () => {
		// The controller element and the layout's own container both carry `dve-print-pages`;
		// zooming both scaled the page by the square of the zoom (150% drew 225%).
		const zoomed = css.match(/\.dve-print-pages\s*\{[^}]*zoom:\s*var\(--dve-zoom/g);
		expect(zoomed).toHaveLength(1);
		expect(css).toMatch(/\.dve-print-pages\s*>\s*\.dve-print-pages\s*\{\s*zoom:\s*1;/);
	});
});
