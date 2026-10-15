/** Widths at or below this collapse the document actions into a More menu. */
export const NARROW_WIDTH = 640;
/** Heights at or below this (phone landscape) let the shell chrome fold away. */
export const SHORT_HEIGHT = 500;

/**
 * Decide how the shell chrome adapts to the visible viewport. Pure so the rules are testable;
 * `mountShell` maps the result onto `data-` attributes the stylesheet reads.
 *
 * - `narrow`: the actions bar becomes one More button.
 * - `drawer`: the app rail becomes an off-canvas drawer (not in a standalone product, which has no rail).
 * - `hideTabs`: a lone open document on a narrow screen needs no tab strip (the brand returns home).
 * - `collapsible`: short editor views may fold the title bar, tabs and actions behind a toggle.
 * - `collapsed`: that fold is currently applied.
 */
export function shellLayout({ width, height, view, documents, teamsOpen, expanded, standalone }) {
	const narrow = width <= NARROW_WIDTH;
	const short = height <= SHORT_HEIGHT;
	const editing = view === 'editor';
	const collapsible = short && editing;
	return {
		narrow,
		drawer: narrow && !standalone,
		hideTabs: narrow && editing && documents === 1 && !teamsOpen,
		collapsible,
		collapsed: collapsible && !expanded,
	};
}
