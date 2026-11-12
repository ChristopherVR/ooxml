/**
 * Page setup the suite's browser checks share, for `addInitScript`. Each function runs inside
 * the page before the app starts, so it must not reference anything outside itself.
 */

/**
 * Skip the start screen and show each app's ribbon. The suite opens documents in Simple view by
 * default, which moves Share, Rename, Download and the other file commands into the More menu
 * and holds the ribbons back; these checks were written against the ribbon layout and press
 * those controls directly. A check that chose a layout itself keeps it.
 * `suite-simple-view-check.mjs` covers the default.
 */
export function ribbonLayoutInit() {
	globalThis.ooxmlSkipStart = true;
	try {
		const key = 'ooxml-suite-settings';
		const stored = JSON.parse(globalThis.localStorage.getItem(key) ?? '{}');
		if (!('layout' in stored))
			globalThis.localStorage.setItem(key, JSON.stringify({ ...stored, layout: 'office' }));
	} catch {
		// Storage is unavailable (a blocked origin): the app falls back to its defaults.
	}
}
