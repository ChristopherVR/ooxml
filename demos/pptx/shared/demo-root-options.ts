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
 *   number; the viewer clamps it into the deck).
 *
 * Only the options present in the query string are returned, so a demo opened
 * without them keeps the binding's own defaults. Kept binding-agnostic on
 * purpose (no import of any viewer package): the value is structurally the
 * bindings' `ViewerRootOptions`.
 */

/** The subset of the root display options given on the query string. */
export interface DemoRootOptions {
	showToolbar?: boolean;
	showThumbnails?: boolean;
	initialSlide?: number;
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
	return options;
}

/** The root display options for the current page (empty when none were given). */
export function currentDemoRootOptions(): DemoRootOptions {
	return typeof window === 'undefined' ? {} : parseDemoRootOptions(window.location.search);
}
