/**
 * Load-time handling of a deck's Google Fonts webfonts, shared by every binding.
 *
 * A family that is neither installed nor embedded is drawn from the Google
 * Fonts fallback (`google-webfonts.ts`). A binding calls
 * {@link prepareDeckWebfonts} as soon as the slides are parsed and awaits it
 * before showing the deck:
 *
 *  1. The deck's fonts are classified without any network request: the
 *     candidates (neither embedded nor installed) that the catalogue serves
 *     become webfonts, the rest are reported as unavailable.
 *  2. The binding's managed `<link>` is injected with the same href its own
 *     webfont effect computes, so that effect later finds it in place and
 *     leaves it alone.
 *  3. Once the stylesheet settles, each face the slides draw (family,
 *     weight, style, and the characters they use, so the right
 *     `unicode-range` subset loads) is loaded through `document.fonts`.
 *
 * The wait ends when every face has loaded or failed, when the stylesheet
 * fails (offline, blocked), or after `timeoutMs` (150 ms, enough for cached
 * faces); the deck is then shown and a face that arrives later swaps in. The
 * download itself is followed for up to `followMs`: progress is reported from
 * `progressDelayMs` on, also after the deck is shown, and once more as `null`
 * when it is over, so the binding can show "Downloading fonts (n of m)" in its
 * status bar and clear it again. A fast or cached load reports nothing.
 *
 * @module render/deck-webfont-readiness
 */
import { hasTextProperties } from 'ooxml-core/pptx';
import type { PptxElement, PptxEmbeddedFont, PptxSlide } from 'ooxml-core/pptx';
import { collectWebfontFaces, type WebfontFace, type WebfontRun } from 'ooxml-core/text';

import {
	WEBFONT_PROGRESS_DELAY_MS,
	WEBFONT_TIMEOUT_MS,
	prefetchWebfonts,
	settlePrefetchedWebfonts,
	waitForFaces,
} from '../../fonts/webfonts';
import {
	isFontFamilyInstalledLocally,
	resolveGoogleWebfontHref,
	selectGoogleWebfontCandidates,
} from './google-webfonts';
import { syncGoogleWebfontStylesheet } from './google-webfonts-link';

/** How long the wait may last before the deck is shown with stand-in faces. */
export const DECK_WEBFONT_TIMEOUT_MS = WEBFONT_TIMEOUT_MS;
/** How long the wait lasts before progress is reported. */
export const DECK_WEBFONT_PROGRESS_DELAY_MS = WEBFONT_PROGRESS_DELAY_MS;
/** How long a download still running when the deck is shown keeps being reported. */
export const DECK_WEBFONT_FOLLOW_MS = 8000;

// Declared here (not aliased to the shared `fonts/webfonts` types) so the bindings' emitted
// declarations name a type that `ooxml-ui/pptx` exports.

/** Progress of the wait, by webfont family (the names the stylesheet serves). */
export interface DeckWebfontProgress {
	/** Families whose faces have all loaded or failed. */
	loaded: number;
	total: number;
	/** Every family being downloaded, in deck order. */
	families: readonly string[];
}

/** Outcome of {@link prepareDeckWebfonts}. */
export interface DeckWebfontReadiness {
	/** The stylesheet the binding's `<link>` now points at, or `null`. */
	href: string | null;
	/** Webfont families the deck waited for (canonical catalogue names). */
	families: readonly string[];
	/** Referenced families with no installed, embedded or catalogue face. */
	unavailable: readonly string[];
	/** The wait ended at `timeoutMs` with faces still downloading. */
	timedOut: boolean;
}

export interface PrepareDeckWebfontsOptions {
	document: Document;
	/** The binding's managed `<link>` id (the one its webfont effect syncs). */
	linkId: string;
	/** The slides the binding's webfont effect will read, so both resolve one href. */
	slides: readonly PptxSlide[];
	embeddedFonts: readonly PptxEmbeddedFont[];
	/** From `resolveGoogleFontsCss2Base`; `null` disables the fallback and the wait. */
	googleFontsBase: string | null;
	/**
	 * Download progress, reported from `progressDelayMs` on (also after the
	 * promise has resolved) and once more as `null` when the download is over.
	 */
	onProgress?: (progress: DeckWebfontProgress | null) => void;
	timeoutMs?: number;
	progressDelayMs?: number;
	followMs?: number;
	isLocallyInstalled?: (family: string) => boolean;
}

function collectRuns(element: PptxElement, runs: WebfontRun[]): void {
	if (hasTextProperties(element)) {
		const base = element.textStyle;
		const segments = element.textSegments ?? [];
		if (segments.length === 0 && base?.fontFamily) {
			runs.push({
				family: base.fontFamily,
				bold: base.bold,
				italic: base.italic,
				text: element.text ?? '',
			});
		}
		for (const segment of segments) {
			runs.push({
				family: segment.style?.fontFamily ?? base?.fontFamily,
				bold: segment.style?.bold ?? base?.bold,
				italic: segment.style?.italic ?? base?.italic,
				text: segment.text ?? '',
			});
		}
	}
	if (element.type === 'group') {
		for (const child of element.children ?? []) {
			collectRuns(child, runs);
		}
	}
	if (element.type === 'table') {
		for (const row of element.tableData?.rows ?? []) {
			for (const cell of row.cells ?? []) {
				for (const run of cell.textRuns ?? []) {
					runs.push({
						family: run.fontFamily ?? cell.style?.fontFamily,
						bold: run.bold ?? cell.style?.bold,
						italic: run.italic ?? cell.style?.italic,
						text: run.text ?? '',
					});
				}
			}
		}
	}
}

/**
 * The faces the slides draw in `candidates`, by served family, and the
 * candidates the catalogue does not serve.
 */
export function collectDeckWebfontFaces(
	slides: readonly PptxSlide[],
	candidates: readonly string[],
): { faces: Map<string, WebfontFace[]>; unavailable: string[] } {
	return collectWebfontFaces(collectDeckWebfontRuns(slides), candidates);
}

/** Every text run the slides draw, for the shared webfont loader. */
export function collectDeckWebfontRuns(slides: readonly PptxSlide[]): WebfontRun[] {
	const runs: WebfontRun[] = [];
	for (const slide of slides) {
		for (const element of slide.elements ?? []) {
			collectRuns(element, runs);
		}
	}
	return runs;
}

/**
 * Start the deck's webfont download from its bytes, before parsing (see `ooxml-ui`'s
 * `prefetchWebfonts`). Each binding calls this with the same `linkId` and base it later passes to
 * {@link prepareDeckWebfonts}. Returns the early href, or `null`; never throws.
 */
export function prefetchDeckWebfonts(options: {
	document: Document;
	linkId: string;
	bytes: Uint8Array | ArrayBuffer;
	googleFontsBase: string | null;
}): string | null {
	return prefetchWebfonts({
		document: options.document,
		linkId: options.linkId,
		bytes: options.bytes,
		base: options.googleFontsBase,
	});
}

/**
 * Inject the deck's webfont stylesheet and wait (bounded) for the faces its
 * slides draw, as described in the module comment. Never rejects.
 */
export async function prepareDeckWebfonts(
	options: PrepareDeckWebfontsOptions,
): Promise<DeckWebfontReadiness> {
	const {
		document: doc,
		linkId,
		slides,
		embeddedFonts,
		googleFontsBase,
		onProgress,
		timeoutMs = DECK_WEBFONT_TIMEOUT_MS,
		progressDelayMs = DECK_WEBFONT_PROGRESS_DELAY_MS,
		followMs = DECK_WEBFONT_FOLLOW_MS,
		isLocallyInstalled = isFontFamilyInstalledLocally,
	} = options;
	settlePrefetchedWebfonts(doc, linkId, googleFontsBase);
	if (googleFontsBase === null) {
		return { href: null, families: [], unavailable: [], timedOut: false };
	}
	const href = await resolveGoogleWebfontHref(
		slides,
		embeddedFonts,
		isLocallyInstalled,
		googleFontsBase,
	);
	const candidates = selectGoogleWebfontCandidates(slides, embeddedFonts, isLocallyInstalled);
	const { faces, unavailable } = collectDeckWebfontFaces(slides, candidates);
	const families = [...faces.keys()];
	const link = syncGoogleWebfontStylesheet(doc, linkId, href);
	const fontSet = doc.fonts as FontFaceSet | undefined;
	if (!link || families.length === 0 || !fontSet || typeof fontSet.load !== 'function') {
		return { href, families, unavailable, timedOut: false };
	}

	return waitForFaces(link, fontSet, faces, families, {
		href,
		unavailable,
		timeoutMs,
		progressDelayMs,
		...(onProgress
			? {
					onProgress,
					followMs: Math.max(followMs, timeoutMs),
					onProgressEnd: () => onProgress(null),
				}
			: {}),
	});
}
