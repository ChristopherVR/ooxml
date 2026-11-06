/**
 * The browser side of the webfont fallback every product shares: whether a family is installed,
 * the managed `<link>` for the CSS2 stylesheet, and a bounded wait for the faces a document draws.
 * Which families to request and how is `ooxml-core/text` (`text/webfonts`); a product reduces its
 * model to `WebfontRun`s and calls {@link prepareWebfonts} between parsing and showing (or
 * relaying out) a document.
 *
 * Moved from PowerPoint's `render/google-webfonts.ts`, `google-webfonts-link.ts` and
 * `deck-webfont-readiness.ts`, which now delegate here.
 *
 * @module fonts/webfonts
 */
import {
	buildGoogleFontsHref,
	collectRunFamilies,
	collectWebfontFaces,
	matchGoogleWebfontFragments,
	selectGoogleWebfontFamilies,
	type WebfontFace,
	type WebfontRun,
} from 'ooxml-core/text';

/**
 * How long a loaded document may wait for its webfonts before it is shown with stand-in faces:
 * about the browser's own block period for `font-display: swap`. Cached faces arrive well inside
 * it; a first download on a slow network does not hold the document back, and a face that arrives
 * later is laid out again in place (each product listens for `document.fonts` `loadingdone`).
 * Waiting longer to avoid one reflow keeps the reader looking at a loading screen.
 */
export const WEBFONT_TIMEOUT_MS = 150;
/** How long the wait lasts before progress is reported. */
export const WEBFONT_PROGRESS_DELAY_MS = 400;

/** Text measured for the local-availability check (mixed glyph widths). */
const INSTALLED_FONT_TEST_STRING = 'mmmmmmmmmmllww';

/**
 * Best-effort "is this family available without any network fetch?". `document.fonts.check`
 * cannot answer it (Chromium reports any unmatched quoted family as loaded), so a test string is
 * measured on a canvas with `<family>, <fallback>` and with `<fallback>` alone: an installed family
 * differs for at least one fallback class. A false negative only loads a face that was not needed.
 */
export function isFontFamilyInstalledLocally(family: string): boolean {
	if (typeof document === 'undefined') return false;
	try {
		const context = document.createElement('canvas').getContext('2d');
		if (!context) return false;
		const quoted = `"${family.replace(/[\\"]/gu, '\\$&')}"`;
		for (const fallback of ['monospace', 'serif', 'sans-serif']) {
			context.font = `72px ${fallback}`;
			const base = context.measureText(INSTALLED_FONT_TEST_STRING).width;
			context.font = `72px ${quoted}, ${fallback}`;
			if (context.measureText(INSTALLED_FONT_TEST_STRING).width !== base) return true;
		}
		return false;
	} catch {
		return false;
	}
}

/** Point the managed `<link id>` at `href`, untouched when it already does; `null` removes it. */
export function syncGoogleWebfontStylesheet(
	doc: Document,
	id: string,
	href: string | null,
): HTMLLinkElement | null {
	const existing = doc.getElementById(id) as HTMLLinkElement | null;
	if (!href) {
		existing?.remove();
		return null;
	}
	// Even assigning the same href can make the browser reload a stylesheet.
	if (existing?.getAttribute('href') === href) return existing;
	const link = existing ?? doc.createElement('link');
	link.id = id;
	link.rel = 'stylesheet';
	link.href = href;
	if (!existing) doc.head.appendChild(link);
	return link;
}

/**
 * Families already resolved this session, by referenced spelling. Once a webfont has loaded it
 * satisfies the canvas check, so re-checking would drop it from the href, remove the very
 * `<link>` that made it available and find it missing again: an oscillation.
 */
const resolvedFamilies = new Set<string>();

/** Reset the session cache (test isolation). */
export function resetWebfontSessionCache(): void {
	resolvedFamilies.clear();
}

/** The families to look up: neither embedded nor installed (and not already resolved). */
export function selectWebfontCandidates(
	families: Iterable<string>,
	embedded: Iterable<string> = [],
	isLocallyInstalled: (family: string) => boolean = isFontFamilyInstalledLocally,
): string[] {
	return selectGoogleWebfontFamilies(
		families,
		embedded,
		(family) => !resolvedFamilies.has(family) && isLocallyInstalled(family),
	);
}

/** The CSS2 fragments for the candidates, recording each match in the session cache. */
export function matchWebfontFragments(candidates: readonly string[]): string[] {
	return matchGoogleWebfontFragments(candidates, (family) => resolvedFamilies.add(family));
}

/** The stylesheet href for the candidates, or `null` (also when `base` is `null`: off). */
export function webfontHref(candidates: readonly string[], base: string | null): string | null {
	if (base === null) return null;
	return buildGoogleFontsHref(matchWebfontFragments(candidates), base);
}

/** Progress of the wait, by webfont family (the names the stylesheet serves). */
export interface WebfontProgress {
	loaded: number;
	total: number;
	families: readonly string[];
}

export interface WebfontReadiness {
	/** The stylesheet the managed `<link>` now points at, or `null`. */
	href: string | null;
	/** Webfont families waited for (canonical catalogue names). */
	families: readonly string[];
	/** Referenced families with no installed, embedded or catalogue face. */
	unavailable: readonly string[];
	/** The wait ended at the timeout with faces still downloading. */
	timedOut: boolean;
}

export interface PrepareWebfontsOptions {
	document: Document;
	/** The product's managed `<link>` id. */
	linkId: string;
	/** The document's text, reduced to runs. */
	runs: readonly WebfontRun[];
	/** Families the document embeds with usable files. */
	embedded?: Iterable<string>;
	/** From `resolveGoogleFontsCss2Base`; `null` turns the fallback and the wait off. */
	base: string | null;
	onProgress?: (progress: WebfontProgress) => void;
	timeoutMs?: number;
	progressDelayMs?: number;
	isLocallyInstalled?: (family: string) => boolean;
}

/** Resolves once `link`'s stylesheet has loaded or failed. */
function stylesheetSettled(link: HTMLLinkElement): Promise<void> {
	if (link.sheet) return Promise.resolve();
	return new Promise((resolve) => {
		link.addEventListener('load', () => resolve(), { once: true });
		link.addEventListener('error', () => resolve(), { once: true });
	});
}

async function loadFace(fonts: FontFaceSet, face: WebfontFace): Promise<void> {
	try {
		await fonts.load(face.font, face.chars.size > 0 ? [...face.chars].join('') : 'x');
	} catch {
		// A face that fails is drawn by the stand-in, as without the wait.
	}
}

/**
 * Inject the stylesheet for the runs' missing families and wait (bounded) for the faces they
 * draw. Never rejects: the wait ends when every face loaded or failed, the stylesheet failed
 * (offline, blocked by policy) or `timeoutMs` passed. Progress is reported only after
 * `progressDelayMs`, so a fast or cached load shows nothing.
 */
export async function prepareWebfonts(options: PrepareWebfontsOptions): Promise<WebfontReadiness> {
	const {
		document: doc,
		linkId,
		runs,
		embedded,
		base,
		onProgress,
		timeoutMs = WEBFONT_TIMEOUT_MS,
		progressDelayMs = WEBFONT_PROGRESS_DELAY_MS,
		isLocallyInstalled = isFontFamilyInstalledLocally,
	} = options;
	if (base === null) return { href: null, families: [], unavailable: [], timedOut: false };
	const candidates = selectWebfontCandidates(
		collectRunFamilies(runs),
		embedded,
		isLocallyInstalled,
	);
	const { faces, unavailable } = collectWebfontFaces(runs, candidates);
	const href = webfontHref(candidates, base);
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
		...(onProgress ? { onProgress } : {}),
	});
}

/**
 * The bounded wait shared with PowerPoint's deck loader.
 *
 * Without `followMs`, a face still downloading when the wait ends reports nothing more (the
 * document is shown and swaps the face in silently). With it, the download keeps being reported
 * after the wait returns, from `progressDelayMs` until every face has loaded or failed, the
 * stylesheet failed, or `followMs` passed; `onProgressEnd` then runs once, when anything was
 * reported, so a host can show the download (a status bar item) and clear it again.
 */
export async function waitForFaces(
	link: HTMLLinkElement,
	fontSet: FontFaceSet,
	faces: ReadonlyMap<string, readonly WebfontFace[]>,
	families: readonly string[],
	context: {
		href: string | null;
		unavailable: readonly string[];
		timeoutMs: number;
		progressDelayMs: number;
		onProgress?: (progress: WebfontProgress) => void;
		followMs?: number;
		onProgressEnd?: () => void;
	},
): Promise<WebfontReadiness> {
	const { href, unavailable, timeoutMs, progressDelayMs, onProgress, followMs, onProgressEnd } =
		context;
	let loaded = 0;
	let reporting = false;
	let over = false;
	const timers: ReturnType<typeof setTimeout>[] = [];
	const report = (): void => {
		if (reporting && !over) onProgress?.({ loaded, total: families.length, families });
	};
	// Ends the reporting for good: when the download settles, at `followMs`, or when the wait
	// returns without `followMs`.
	const end = (): void => {
		if (over) return;
		over = true;
		for (const timer of timers) clearTimeout(timer);
		if (reporting) {
			reporting = false;
			onProgressEnd?.();
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
	timers.push(
		setTimeout(() => {
			reporting = true;
			report();
		}, progressDelayMs),
	);
	if (followMs !== undefined) {
		timers.push(setTimeout(end, followMs));
		void download.then(end);
	}
	let waitTimer: ReturnType<typeof setTimeout> | undefined;
	const timedOut = await Promise.race([
		download.then(() => false),
		new Promise<boolean>((resolve) => {
			waitTimer = setTimeout(() => resolve(true), timeoutMs);
		}),
	]);
	clearTimeout(waitTimer);
	// Without `followMs` a face still downloading must not report into a document already shown.
	if (followMs === undefined || !timedOut) end();
	return { href, families, unavailable, timedOut };
}
