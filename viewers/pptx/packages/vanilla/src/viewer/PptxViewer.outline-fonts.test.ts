import { afterEach, describe, expect, it } from 'vitest';

import { embeddedFontList, retainedFontLists } from '../../../shared/test-utils/outline-fonts.mjs';
import { glyphOutlineFontCache } from './glyph-outline-cache';
import { createPptxViewer } from './PptxViewer';
import type { PptxViewer } from './PptxViewer';

afterEach(() => {
	document.body.replaceChildren();
});

describe('embedded fonts in the glyph-outline cache', () => {
	it('are released when the deck is replaced and when the viewer is destroyed', () => {
		const container = document.createElement('div');
		document.body.appendChild(container);
		const viewer = createPptxViewer(container) as PptxViewer;
		const first = embeddedFontList();
		const second = embeddedFontList();
		try {
			viewer.store.set({ embeddedFonts: first });
			expect(retainedFontLists(glyphOutlineFontCache).has(first)).toBeTruthy();
			// A state change that keeps the font list keeps its registration.
			viewer.store.set({ slides: [] });
			expect(retainedFontLists(glyphOutlineFontCache).has(first)).toBeTruthy();

			viewer.store.set({ embeddedFonts: second });
			const retained = retainedFontLists(glyphOutlineFontCache);
			expect(retained.has(first)).toBeFalsy();
			expect(retained.has(second)).toBeTruthy();
		} finally {
			viewer.destroy();
		}
		const retained = retainedFontLists(glyphOutlineFontCache);
		expect(retained.has(first)).toBeFalsy();
		expect(retained.has(second)).toBeFalsy();
	});
});
