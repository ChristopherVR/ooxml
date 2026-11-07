/**
 * Glyph-outline bytes for a deck's Google Fonts webfonts, shared by every
 * binding (each used to carry its own copy of this loop).
 *
 * Only glyph-envelope WordArt (inflate, deflate, can, the waves and rings,
 * see `hasGlyphEnvelope`) draws real glyph outlines, and it falls back to
 * tracing the browser's own rendering without them. So the font files are
 * fetched only for families that such text uses, rather than for every
 * webfont of every deck, and each stylesheet and file is fetched once (see
 * `text-warp-outline-webfont-fetch.ts`).
 *
 * @module render/deck-webfont-outlines
 */
import { hasTextProperties } from 'ooxml-core/pptx';
import type { PptxElement, PptxEmbeddedFont, PptxSlide } from 'ooxml-core/pptx';

import { selectTrustedEmbeddedFonts } from './embedded-font-trust';
import { resolveGoogleWebfontHref, selectGoogleWebfontFamilies } from './google-webfonts';
import { hasGlyphEnvelope } from './text-warp-envelope-curves';
import { fetchGoogleWebfontOutlineBytes } from './text-warp-outline-webfont-fetch';
import type { FetchLike } from './text-warp-outline-webfont-fetch';
import { collectFontsFromElement } from './used-fonts';

/** What {@link loadDeckWebfontOutlines} needs from a `GlyphOutlineFontCache`. */
export interface WebfontOutlineRegistrar {
	registerFontBytes(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): boolean;
}

function collectEnvelopeFonts(element: PptxElement, fonts: Set<string>): void {
	if (hasTextProperties(element) && hasGlyphEnvelope(element.textStyle?.textWarpPreset)) {
		collectFontsFromElement(element, fonts);
	}
	if (element.type === 'group') {
		for (const child of element.children ?? []) {
			collectEnvelopeFonts(child, fonts);
		}
	}
}

/** Families used by glyph-envelope WordArt anywhere in `slides`. */
export function collectGlyphEnvelopeFontFamilies(slides: readonly PptxSlide[]): Set<string> {
	const fonts = new Set<string>();
	for (const slide of slides) {
		for (const element of slide.elements ?? []) {
			collectEnvelopeFonts(element, fonts);
		}
	}
	return fonts;
}

export interface LoadDeckWebfontOutlinesOptions {
	slides: readonly PptxSlide[];
	embeddedFonts: readonly PptxEmbeddedFont[];
	/** From `resolveGoogleFontsCss2Base`; `null` disables the fallback. */
	googleFontsBase: string | null;
	cache: WebfontOutlineRegistrar;
	/** Checked between fetches; a superseded deck stops registering. */
	isCancelled?: () => boolean;
	fetchImpl?: FetchLike;
}

const VARIANTS: ReadonlyArray<{ bold: boolean; italic: boolean }> = [
	{ bold: false, italic: false },
	{ bold: true, italic: false },
];

/**
 * Fetch and register the outline bytes of the webfonts glyph-envelope
 * WordArt uses. Resolves `true` when any font was registered (the binding
 * then repaints), `false` otherwise. Best-effort; never rejects.
 */
export async function loadDeckWebfontOutlines(
	options: LoadDeckWebfontOutlinesOptions,
): Promise<boolean> {
	const { slides, embeddedFonts, googleFontsBase, cache, isCancelled = () => false } = options;
	const fetchImpl =
		options.fetchImpl ??
		(typeof fetch === 'function' ? (fetch as unknown as FetchLike) : undefined);
	if (!fetchImpl) {
		return false;
	}
	const families = selectGoogleWebfontFamilies(
		collectGlyphEnvelopeFontFamilies(slides),
		selectTrustedEmbeddedFonts(embeddedFonts).map((font) => font.name),
	);
	if (families.length === 0) {
		return false;
	}
	const href = await resolveGoogleWebfontHref(slides, embeddedFonts, undefined, googleFontsBase);
	if (!href || isCancelled()) {
		return false;
	}
	let registeredAny = false;
	for (const family of families) {
		for (const { bold, italic } of VARIANTS) {
			if (isCancelled()) {
				return false;
			}
			const bytes = await fetchGoogleWebfontOutlineBytes(href, family, bold, italic, fetchImpl);
			if (bytes && !isCancelled() && cache.registerFontBytes(family, bold, italic, bytes)) {
				registeredAny = true;
			}
		}
	}
	return registeredAny && !isCancelled();
}
