/**
 * Load-time wait for a deck's Google Fonts webfonts, shared by every binding.
 *
 * A family that is neither installed nor embedded is drawn from the Google
 * Fonts fallback (`google-webfonts.ts`). Injected only after the deck
 * renders, its stylesheet arrives a moment later and every slide visibly
 * swaps from the browser's stand-in face to the real one. Instead, a binding
 * calls {@link prepareDeckWebfonts} between parsing a deck and showing it,
 * while its loading state is still up:
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
 * fails (offline, blocked), or after `timeoutMs`; the deck is then shown and
 * a face that arrives later swaps in as before. Progress is reported only
 * once the wait has lasted `progressDelayMs`, so a fast or cached load shows
 * no extra message.
 *
 * @module render/deck-webfont-readiness
 */
import { hasTextProperties } from 'ooxml-core/pptx';
import type { PptxElement, PptxEmbeddedFont, PptxSlide } from 'ooxml-core/pptx';

import { resolveFontItalic, resolveFontWeight, resolveGoogleFontsFace } from './font-weight-names';
import {
	isFontFamilyInstalledLocally,
	resolveGoogleWebfontHref,
	selectGoogleWebfontCandidates,
} from './google-webfonts';
import { syncGoogleWebfontStylesheet } from './google-webfonts-link';
import { findMetricCompatibleGoogleFontsFamily } from './google-webfonts-metric-clones';

/** How long the wait may last before the deck is shown with stand-in faces. */
export const DECK_WEBFONT_TIMEOUT_MS = 8000;
/** How long the wait lasts before progress is reported. */
export const DECK_WEBFONT_PROGRESS_DELAY_MS = 400;

/** Characters collected per face for its `unicode-range` subset (enough for any script mix). */
const MAX_SAMPLE_CHARS = 256;

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
	onProgress?: (progress: DeckWebfontProgress) => void;
	timeoutMs?: number;
	progressDelayMs?: number;
	isLocallyInstalled?: (family: string) => boolean;
}

/** One face the slides draw: CSS `font` shorthand plus the characters drawn in it. */
interface FaceUse {
	family: string;
	font: string;
	chars: Set<string>;
}

interface RunUse {
	family: string | undefined;
	bold: boolean | undefined;
	italic: boolean | undefined;
	text: string;
}

function collectRuns(element: PptxElement, runs: RunUse[]): void {
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

/** The family the stylesheet serves for `family`, or `null` when the catalogue has none. */
function servedFamily(family: string): string | null {
	return resolveGoogleFontsFace(family)?.family ?? findMetricCompatibleGoogleFontsFamily(family);
}

function quoteFamily(family: string): string {
	return `"${family.replace(/[\\"]/gu, '\\$&')}"`;
}

/**
 * The faces the slides draw in `candidates`, by served family, and the
 * candidates the catalogue does not serve.
 */
export function collectDeckWebfontFaces(
	slides: readonly PptxSlide[],
	candidates: readonly string[],
): { faces: Map<string, FaceUse[]>; unavailable: string[] } {
	const wanted = new Set(candidates);
	const unavailable = candidates.filter((family) => servedFamily(family) === null);
	const runs: RunUse[] = [];
	for (const slide of slides) {
		for (const element of slide.elements ?? []) {
			collectRuns(element, runs);
		}
	}
	const faces = new Map<string, FaceUse[]>();
	const byFont = new Map<string, FaceUse>();
	for (const run of runs) {
		if (!run.family || !wanted.has(run.family)) {
			continue;
		}
		const served = servedFamily(run.family);
		if (served === null) {
			continue;
		}
		const weight = resolveFontWeight(run.family, run.bold);
		const style = resolveFontItalic(run.family, run.italic) ? 'italic ' : '';
		const font = `${style}${weight} 16px ${quoteFamily(served)}`;
		let face = byFont.get(font);
		if (!face) {
			face = { family: served, font, chars: new Set() };
			byFont.set(font, face);
			const list = faces.get(served);
			if (list) {
				list.push(face);
			} else {
				faces.set(served, [face]);
			}
		}
		for (const char of run.text) {
			if (face.chars.size >= MAX_SAMPLE_CHARS) {
				break;
			}
			if (char.trim() !== '') {
				face.chars.add(char);
			}
		}
	}
	return { faces, unavailable };
}

/** Resolves once `link`'s stylesheet has loaded or failed. */
function stylesheetSettled(link: HTMLLinkElement): Promise<void> {
	if (link.sheet) {
		return Promise.resolve();
	}
	return new Promise((resolve) => {
		link.addEventListener('load', () => resolve(), { once: true });
		link.addEventListener('error', () => resolve(), { once: true });
	});
}

async function loadFace(fonts: FontFaceSet, face: FaceUse): Promise<void> {
	try {
		await fonts.load(face.font, face.chars.size > 0 ? [...face.chars].join('') : 'x');
	} catch {
		// A face that fails is drawn by the stand-in, as without the wait.
	}
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
		isLocallyInstalled = isFontFamilyInstalledLocally,
	} = options;
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

	let loaded = 0;
	let reporting = false;
	const report = (): void => {
		if (reporting) {
			onProgress?.({ loaded, total: families.length, families });
		}
	};
	const download = (async () => {
		await stylesheetSettled(link);
		await Promise.all(
			[...faces.values()].map(async (familyFaces) => {
				await Promise.all(familyFaces.map((face) => loadFace(fontSet, face)));
				loaded += 1;
				report();
			}),
		);
	})();

	let finished = false;
	const timers: ReturnType<typeof setTimeout>[] = [];
	const timedOut = await Promise.race([
		download.then(() => false),
		new Promise<boolean>((resolve) => {
			timers.push(setTimeout(() => resolve(true), timeoutMs));
			timers.push(
				setTimeout(() => {
					if (!finished) {
						reporting = true;
						report();
					}
				}, progressDelayMs),
			);
		}),
	]);
	// A face still downloading after a timeout must not report into a deck already shown.
	finished = true;
	reporting = false;
	for (const timer of timers) {
		clearTimeout(timer);
	}
	return { href, families, unavailable, timedOut };
}
