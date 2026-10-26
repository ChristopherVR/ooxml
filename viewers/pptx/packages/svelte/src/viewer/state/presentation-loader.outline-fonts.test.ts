import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	embedFontsOnLoad,
	retainedFontLists,
} from '../../../../shared/test-utils/outline-fonts.mjs';
import { presentationLoadFixtures } from '../../../../shared/test-utils/presentation-load.mjs';
import { glyphOutlineFontCache } from './glyph-outline-cache.svelte';
import { PresentationLoader } from './presentation-loader.svelte';

afterEach(() => vi.restoreAllMocks());

describe('embedded fonts in the glyph-outline cache', () => {
	it('are released when the deck is replaced and when the loader is disposed', async () => {
		const fixtures = await presentationLoadFixtures(PptxHandler);
		const lists = embedFontsOnLoad(PptxHandler);
		const loader = new PresentationLoader();
		try {
			await loader.load(fixtures.first);
			expect(loader.embeddedFonts).toBe(lists[0]);
			expect(retainedFontLists(glyphOutlineFontCache).has(lists[0]!)).toBeTruthy();

			await loader.load(fixtures.second);
			const retained = retainedFontLists(glyphOutlineFontCache);
			expect(retained.has(lists[0]!)).toBeFalsy();
			expect(retained.has(lists[1]!)).toBeTruthy();
		} finally {
			loader.dispose();
		}
		const retained = retainedFontLists(glyphOutlineFontCache);
		expect(lists).toHaveLength(2);
		expect(lists.some((list) => retained.has(list))).toBeFalsy();
	}, 30_000);
});
