import { describe, expect, it } from 'vitest';

import { buildViewerCss } from './css';

describe('protected view bar', () => {
	it('stays hidden on a host page without a CSS reset', () => {
		// The bar is always in the DOM and toggles `hidden`; its `display: flex` rule must not win.
		expect(buildViewerCss()).toMatch(/\.pptxv-protected-view\[hidden\] \{ display: none; \}/u);
	});
});
