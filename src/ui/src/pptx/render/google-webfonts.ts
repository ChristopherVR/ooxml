/**
 * google-webfonts.ts: Google Fonts webfont fallback for referenced font
 * families, shared by every binding.
 *
 * A deck may reference a font family that is neither installed on the
 * reader's machine nor embedded in the .pptx itself. PowerPoint still renders
 * such decks correctly because Microsoft 365 silently downloads its "cloud
 * fonts" on demand; a browser has no equivalent mechanism. For referenced
 * families that the Google Fonts CSS2 API serves, each binding injects a
 * `<link rel="stylesheet">` so the text renders with the intended face
 * anyway.
 *
 * Which families the API serves is answered by the bundled catalogue
 * (`google-fonts-catalogue.ts`, regenerated from Google's own metadata feed
 * by `bun run fonts:catalogue`), never by probing the API per family: a
 * probe of an unknown family answers 400 without CORS headers, which the
 * browser reports as an uncatchable console error, and it discloses every
 * missing family name to Google. With the catalogue the only request ever
 * made is the stylesheet for families already known to be served. The API is
 * lenient about weights and styles (it serves only the ones a family has),
 * so one universal axis spec is safe for every catalogue family.
 *
 * Families detected as locally installed are dropped before the href is
 * built, so an installed face is used as-is. Everything here is pure; the DOM
 * side effect (injecting / updating / removing the managed `<link>` element)
 * stays in each binding, and the element id is binding-specific.
 *
 * A referenced family that the API does not serve under its OWN name (e.g.
 * "Calibri", which Microsoft never published to Google Fonts) is not simply
 * dropped: `findMetricCompatibleGoogleFontsFamily` looks it up in a small,
 * curated map of VERIFIED metric-compatible clones (the same names
 * `ooxml-core/pptx`'s `font-substitution.ts` already puts second in the CSS
 * `font-family` list, e.g. Calibri -> Carlito) and requests that clone
 * instead when the catalogue serves it. Because that family is already
 * second in the CSS chain, once its stylesheet loads the browser measures
 * text with matching advance widths without any binding-specific wiring: the
 * fix lives once, here, and every binding's existing
 * `resolveGoogleWebfontHref` call picks it up. Some metric-compatible fonts
 * (Segoe UI's, Selawik) are not on Google Fonts at all; for those
 * `googleFontsBaseUrl` can be pointed at a self-hosted mirror of the CSS2 API
 * (see {@link resolveGoogleWebfontHref}), otherwise nothing is fetched and
 * the CSS chain falls through to its next, non-metric-matched entry as it
 * already did.
 */

import type { PptxEmbeddedFont, PptxSlide } from 'ooxml-core/pptx';
import { GOOGLE_FONTS_CSS2_BASE } from 'ooxml-core/text';

import {
	isFontFamilyInstalledLocally,
	matchWebfontFragments,
	resetWebfontSessionCache,
	selectWebfontCandidates,
	webfontHref,
} from '../../fonts/webfonts';
import { selectTrustedEmbeddedFonts } from './embedded-font-trust';
import { collectFontsFromElement } from './used-fonts';

// The format-neutral request and the browser loader moved to `ooxml-core/text` and
// `ooxml-ui`'s `fonts/webfonts.ts` so Word, Excel and Visio share them; these stay exported here.
export {
	GOOGLE_FONTS_CSS2_BASE,
	GOOGLE_WEBFONT_AXIS,
	buildGoogleFontsFragment,
	buildGoogleFontsHref,
	findGoogleFontsFamily,
	findMetricCompatibleGoogleFontsFamily,
	normalizeGoogleFontsBaseUrl,
	resolveGoogleFontsCss2Base,
	selectGoogleWebfontFamilies,
	type GoogleFontsOption,
} from 'ooxml-core/text';
export { isFontFamilyInstalledLocally } from '../../fonts/webfonts';

/**
 * Collect all unique font family names referenced across slide elements,
 * recursing into group children.
 *
 * Delegates to the same walker the Fonts inspector uses, so the element-level
 * `textStyle.fontFamily` counts too: segment-less text (a title whose face is
 * set once on the body rather than per run) used to be skipped here, so its
 * webfont was never requested and it painted in the fallback face.
 */
export function collectReferencedFontFamilies(slides: readonly PptxSlide[]): Set<string> {
	const families = new Set<string>();
	for (const slide of slides) {
		for (const el of slide.elements) {
			collectFontsFromElement(el, families);
		}
	}
	return families;
}

/** Reset the session cache (test isolation). */
export function resetGoogleWebfontSessionCache(): void {
	resetWebfontSessionCache();
}

/**
 * The query fragments for the candidate families the catalogue knows (see
 * `ooxml-core/text`'s `matchGoogleWebfontFragments`), recording each match in
 * the shared session cache so a loaded webfont is not mistaken for an
 * installed one on the next call.
 */
export function matchGoogleWebfontFragments(families: readonly string[]): string[] {
	return matchWebfontFragments(families);
}

/**
 * One-stop helper the bindings call from their reactive wiring: resolve the
 * href for a loaded deck's slides + embedded fonts (`null` when no stylesheet
 * is needed). Families available locally are used as-is; the rest are matched
 * against the bundled catalogue. An embedded family only counts as available
 * when its files match it (see {@link selectTrustedEmbeddedFonts}). The
 * local-install check is skipped for families the session already resolved.
 *
 * @param googleFontsBaseUrl - Overrides `fonts.googleapis.com` for a
 *   self-hosted CSS2 API mirror. `null` (see `resolveGoogleFontsCss2Base`)
 *   disables the fallback: the result is always `null`.
 */
export async function resolveGoogleWebfontHref(
	slides: readonly PptxSlide[],
	embeddedFonts: readonly PptxEmbeddedFont[],
	isLocallyInstalled: (family: string) => boolean = isFontFamilyInstalledLocally,
	googleFontsBaseUrl?: string | null,
): Promise<string | null> {
	if (googleFontsBaseUrl === null) {
		return null;
	}
	const candidates = selectGoogleWebfontCandidates(slides, embeddedFonts, isLocallyInstalled);
	return webfontHref(candidates, googleFontsBaseUrl ?? GOOGLE_FONTS_CSS2_BASE);
}

/**
 * The referenced families {@link resolveGoogleWebfontHref} looks up in the
 * catalogue: neither embedded (with matching files) nor installed. Shared so
 * the load-time wait (`deck-webfont-readiness.ts`) waits for exactly the
 * families the stylesheet serves.
 */
export function selectGoogleWebfontCandidates(
	slides: readonly PptxSlide[],
	embeddedFonts: readonly PptxEmbeddedFont[],
	isLocallyInstalled: (family: string) => boolean = isFontFamilyInstalledLocally,
): string[] {
	return selectWebfontCandidates(
		collectReferencedFontFamilies(slides),
		selectTrustedEmbeddedFonts(embeddedFonts).map((font) => font.name),
		isLocallyInstalled,
	);
}
