import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { beforeEach, describe, expect, it, vi } from 'vitest';

import { collectGlyphEnvelopeFontFamilies, loadDeckWebfontOutlines } from './deck-webfont-outlines';
import { resetGoogleWebfontSessionCache } from './google-webfonts';
import type { FetchLike } from './text-warp-outline-webfont-fetch';

function textEl(fontFamily: string, textWarpPreset?: string): PptxElement {
	return {
		type: 'text',
		textStyle: { fontFamily, textWarpPreset },
		textSegments: [{ text: 'Hi', style: {} }],
	} as unknown as PptxElement;
}

function slide(...elements: PptxElement[]): PptxSlide {
	return { elements } as unknown as PptxSlide;
}

const CSS = `@font-face {
  font-family: 'Raleway';
  font-style: normal;
  font-weight: 400;
  src: url(https://fonts.gstatic.com/raleway-400.woff2) format('woff2');
}
@font-face {
  font-family: 'Raleway';
  font-style: normal;
  font-weight: 700;
  src: url(https://fonts.gstatic.com/raleway-700.woff2) format('woff2');
}`;

function fakeFetch(): FetchLike {
	return vi.fn(async (url: string) => ({
		ok: true,
		text: async () => CSS,
		arrayBuffer: async () => new Uint8Array([0, 1, 0, 0, url.length]).buffer,
	}));
}

beforeEach(() => {
	resetGoogleWebfontSessionCache();
});

describe('collectGlyphEnvelopeFontFamilies', () => {
	it('collects only the families of glyph-envelope WordArt, inside groups too', () => {
		const group = {
			type: 'group',
			children: [textEl('Lobster', 'textCanUp')],
		} as unknown as PptxElement;
		const families = collectGlyphEnvelopeFontFamilies([
			slide(textEl('Raleway'), textEl('Playfair Display', 'textArchUp'), group),
		]);
		expect([...families]).toEqual(['Lobster']);
	});
});

describe('loadDeckWebfontOutlines', () => {
	const base = 'https://fonts.googleapis.com/css2';

	it('fetches nothing for a deck without glyph-envelope WordArt', async () => {
		const fetchImpl = fakeFetch();
		const cache = { registerFontBytes: vi.fn(() => true) };
		const registered = await loadDeckWebfontOutlines({
			slides: [slide(textEl('Raleway'), textEl('Playfair Display'))],
			embeddedFonts: [],
			googleFontsBase: base,
			cache,
			fetchImpl,
		});
		expect(registered).toBe(false);
		expect(fetchImpl).not.toHaveBeenCalled();
	});

	it('registers regular and bold outlines, fetching the stylesheet once', async () => {
		const fetchImpl = fakeFetch();
		const cache = { registerFontBytes: vi.fn(() => true) };
		const registered = await loadDeckWebfontOutlines({
			slides: [slide(textEl('Raleway', 'textInflate'))],
			embeddedFonts: [],
			googleFontsBase: base,
			cache,
			fetchImpl,
		});
		expect(registered).toBe(true);
		expect(cache.registerFontBytes.mock.calls.map((call) => call.slice(0, 3))).toEqual([
			['Raleway', false, false],
			['Raleway', true, false],
		]);
		const urls = vi.mocked(fetchImpl).mock.calls.map(([url]) => url);
		expect(urls.filter((url) => url.startsWith(base))).toHaveLength(1);
	});

	it('stops registering once cancelled', async () => {
		const fetchImpl = fakeFetch();
		const cache = { registerFontBytes: vi.fn(() => true) };
		const registered = await loadDeckWebfontOutlines({
			slides: [slide(textEl('Raleway', 'textInflate'))],
			embeddedFonts: [],
			googleFontsBase: base,
			cache,
			fetchImpl,
			isCancelled: () => true,
		});
		expect(registered).toBe(false);
		expect(cache.registerFontBytes).not.toHaveBeenCalled();
	});
});
