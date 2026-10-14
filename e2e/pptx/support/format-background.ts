/**
 * Which bindings open Design > Format Background as a docked panel.
 *
 * VanillaJS and Svelte open a docked panel (a `dialog` named "Format
 * Background"); React, Vue and Angular open the format pane instead. A spec
 * about the docked panel must assert that it opened where it exists, rather
 * than pass vacuously against a dialog that is missing, so it asks here.
 *
 * Kept out of the specs because `scripts/check-e2e-neutrality.mjs` forbids
 * branching on the Playwright project inside them.
 *
 * @module e2e/support/format-background
 */

const DOCKED_PANEL_BINDINGS: ReadonlySet<string> = new Set(['vanilla', 'svelte']);

/**
 * Why this binding cannot run a docked Format Background check, or `null`
 * when it opens the docked panel and the check must run.
 */
export function dockedFormatBackgroundGap(binding: string): string | null {
	return DOCKED_PANEL_BINDINGS.has(binding)
		? null
		: 'this binding opens Format Background in the format pane, not a docked panel';
}
