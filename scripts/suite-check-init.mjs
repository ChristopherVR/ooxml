/**
 * Page setup the suite's browser checks share, for `addInitScript`. Each init function runs inside
 * the page before the app starts, so it must not reference anything outside itself. `fileCommand`
 * is a helper for the checks themselves and runs in Node.
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

/**
 * Runs one of the document's file commands the way a user does: from the bar when it is there,
 * else from the More menu (the ribbon layout and phones keep them in the menu).
 */
export async function fileCommand(page, id) {
	const button = page.locator(`#${id}`);
	if (await button.isVisible()) return button.click();
	const label = (await button.textContent()).trim();
	await page.locator('#actions-more').click();
	await page.locator('#actions-menu [role="menuitem"]').filter({ hasText: label }).first().click();
}
