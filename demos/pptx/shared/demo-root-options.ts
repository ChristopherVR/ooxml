/**
 * Demo-only plumbing for the root display options every PowerPoint binding
 * accepts (`showToolbar`, `showThumbnails`, `initialSlide`; the defaults and
 * clamping live in `src/ui/src/pptx/render/viewer-root-options.ts`). Shared by
 * all five demos so the framework-neutral e2e spec
 * (`e2e/pptx/root-display-options.spec.ts`) drives every binding the same way:
 *
 * - `?showToolbar=0` / `?showThumbnails=0` turn the chrome or the slide pane off
 *   (`1` / `true` turn them on; any other value is ignored),
 * - `?initialSlide=<n>` opens the deck on zero-based slide `n` (any finite
 *   number; the viewer clamps it into the deck),
 * - `?fitPadding=<px>` sets the per-side fit padding (any finite number of at
 *   least 0), so a spec can fit the slide flush against the viewport,
 * - `?googleFonts=0` turns the Google Fonts fallback off, and
 *   `?googleFontsBaseUrl=<url>` points it at a self-hosted CSS2 mirror,
 * - `?allowExternalContent=0` / `?allowExternalContent=1` set the host
 *   external-content switch (absent leaves it to the Trust Center option),
 * - `?listConversionScope=list` / `?listConversionScope=paragraph` set how far
 *   a bare caret's Bullets / Numbering reaches (absent keeps the PowerPoint
 *   default, `'paragraph'`),
 * - `?maxInputBytes=`, `?maxUncompressedBytes=`, `?maxEntries=`,
 *   `?maxSlides=` and `?maxShapes=` set the matching field of the host
 *   `limits` option (a positive number; any other value is ignored).
 *
 * Only the options present in the query string are returned, so a demo opened
 * without them keeps the binding's own defaults. Kept binding-agnostic on
 * purpose (no import of any viewer package): the value is structurally the
 * bindings' `ViewerRootOptions`.
 */

/** The subset of the root display options given on the query string. */
export interface DemoRootOptions {
	inspectorInitiallyOpen?: boolean;
	showToolbar?: boolean;
	showThumbnails?: boolean;
	initialSlide?: number;
	fitPadding?: number;
	googleFonts?: false | { baseUrl: string };
	allowExternalContent?: boolean;
	listConversionScope?: 'paragraph' | 'list';
	limits?: DemoLoadLimits;
}

/** The `limits` fields a demo URL can set (structurally the core's `PptxLoadLimits`). */
export interface DemoLoadLimits {
	maxInputBytes?: number;
	maxUncompressedBytes?: number;
	maxEntries?: number;
	maxSlides?: number;
	maxShapes?: number;
}

const LIMIT_PARAMS = [
	'maxInputBytes',
	'maxUncompressedBytes',
	'maxEntries',
	'maxSlides',
	'maxShapes',
] as const satisfies readonly (keyof DemoLoadLimits)[];

function parseLimits(params: URLSearchParams): DemoLoadLimits | undefined {
	const limits: DemoLoadLimits = {};
	let found = false;
	for (const name of LIMIT_PARAMS) {
		const raw = params.get(name);
		const value = raw === null || raw.trim() === '' ? Number.NaN : Number(raw);
		if (Number.isFinite(value) && value > 0) {
			limits[name] = value;
			found = true;
		}
	}
	return found ? limits : undefined;
}

function parseFlag(raw: string | null): boolean | undefined {
	if (raw === '0' || raw === 'false') {
		return false;
	}
	if (raw === '1' || raw === 'true') {
		return true;
	}
	return undefined;
}

/** Read the root display options from `search` (the page query string). */
export function parseDemoRootOptions(search: string): DemoRootOptions {
	const params = new URLSearchParams(search);
	const options: DemoRootOptions = {};
	const inspectorInitiallyOpen = parseFlag(params.get('inspectorInitiallyOpen'));
	if (inspectorInitiallyOpen !== undefined) options.inspectorInitiallyOpen = inspectorInitiallyOpen;
	const showToolbar = parseFlag(params.get('showToolbar'));
	if (showToolbar !== undefined) {
		options.showToolbar = showToolbar;
	}
	const showThumbnails = parseFlag(params.get('showThumbnails'));
	if (showThumbnails !== undefined) {
		options.showThumbnails = showThumbnails;
	}
	const rawSlide = params.get('initialSlide');
	const initialSlide = rawSlide === null || rawSlide.trim() === '' ? Number.NaN : Number(rawSlide);
	if (Number.isFinite(initialSlide)) {
		options.initialSlide = initialSlide;
	}
	const rawPadding = params.get('fitPadding');
	const fitPadding =
		rawPadding === null || rawPadding.trim() === '' ? Number.NaN : Number(rawPadding);
	if (Number.isFinite(fitPadding) && fitPadding >= 0) {
		options.fitPadding = fitPadding;
	}
	const googleFontsBaseUrl = params.get('googleFontsBaseUrl');
	if (googleFontsBaseUrl) {
		options.googleFonts = { baseUrl: googleFontsBaseUrl };
	} else if (parseFlag(params.get('googleFonts')) === false) {
		options.googleFonts = false;
	}
	const allowExternalContent = parseFlag(params.get('allowExternalContent'));
	if (allowExternalContent !== undefined) {
		options.allowExternalContent = allowExternalContent;
	}
	const listConversionScope = params.get('listConversionScope');
	if (listConversionScope === 'paragraph' || listConversionScope === 'list') {
		options.listConversionScope = listConversionScope;
	}
	const limits = parseLimits(params);
	if (limits) {
		options.limits = limits;
	}
	return options;
}

/** The root display options for the current page (empty when none were given). */
export function currentDemoRootOptions(): DemoRootOptions {
	return typeof window === 'undefined' ? {} : parseDemoRootOptions(window.location.search);
}
