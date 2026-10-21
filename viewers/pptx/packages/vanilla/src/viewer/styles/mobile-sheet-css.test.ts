import { describe, expect, it } from 'vitest';

import { buildViewerCss } from './css';
import { MOBILE_SHEET_CSS } from './mobile-sheet-css';

/** The z-index declared in the first rule whose selector is exactly `selector`. */
function zOf(css: string, selector: string): number {
	const start = css.indexOf(`${selector} {`);
	return Number(/z-index: (\d+)/.exec(css.slice(start, css.indexOf('}', start)))?.[1]);
}

describe('phone sheet stacking and layout', () => {
	it('stacks the sheet and bar above the selection overlay (z-index 58) and its tools (61)', () => {
		expect(zOf(MOBILE_SHEET_CSS, '.pptxv-mobile-sheet-host')).toBeGreaterThan(61);
		expect(zOf(MOBILE_SHEET_CSS, '.pptxv-mobile-actions > pptx-ui-mobile-bar')).toBeGreaterThan(61);
	});

	it('isolates the viewer so its layers cannot paint over the host page', () => {
		const rule = [...buildViewerCss().matchAll(/\n\.pptxv \{[^}]*\}/gu)]
			.map(([css]) => css)
			.join('\n');
		expect(rule).toContain('isolation: isolate');
	});

	it('lays the Format sheet out as one column with a compact three-up tab row', () => {
		expect(MOBILE_SHEET_CSS).toContain('.pptxv-mobile-sheet .pptxv-inspector { display: block;');
		expect(MOBILE_SHEET_CSS).toContain('grid-template-columns: repeat(3, minmax(0, 1fr))');
	});

	it('gives sheet content the full row so the comment box is not half width', () => {
		expect(MOBILE_SHEET_CSS).toContain('.pptxv-mobile-comment-add');
		expect(MOBILE_SHEET_CSS).toContain('flex: 1 1 100%');
	});
});
