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
 *   `?googleFontsBaseUrl=<url>` points it at a self-hosted CSS2 mirror.
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
	return options;
}

/** The root display options for the current page (empty when none were given). */
export function currentDemoRootOptions(): DemoRootOptions {
	return typeof window === 'undefined' ? {} : parseDemoRootOptions(window.location.search);
}
