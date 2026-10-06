/**
 * The suite's tab model: a pure list of tabs and which one is active, with no DOM and no host
 * assumptions, so the web shell and the desktop shell drive the same code. Every tab is one app
 * (and the framework demo or editor it runs on); the same app may be open in several tabs. Home is
 * not a tab: it is the state where nothing is active.
 */

export interface SuiteTab {
	readonly id: string;
	readonly app: string;
	readonly framework: string;
}

export interface SuiteTabState {
	readonly tabs: readonly SuiteTab[];
	/** The active tab's id, or null for Home. */
	readonly active: string | null;
}

export const EMPTY_SUITE_STATE: SuiteTabState = { tabs: [], active: null };

/** Read a persisted state, dropping anything malformed, duplicated or pointing at a missing tab. */
export function parseSuiteState(value: unknown): SuiteTabState {
	if (!value || typeof value !== 'object') return EMPTY_SUITE_STATE;
	const raw = value as { tabs?: unknown; active?: unknown };
	if (!Array.isArray(raw.tabs)) return EMPTY_SUITE_STATE;
	const tabs: SuiteTab[] = [];
	for (const item of raw.tabs as unknown[]) {
		if (!item || typeof item !== 'object') continue;
		const { id, app, framework } = item as Record<string, unknown>;
		if (typeof id !== 'string' || typeof app !== 'string' || typeof framework !== 'string')
			continue;
		if (tabs.some((tab) => tab.id === id)) continue;
		tabs.push({ id, app, framework });
	}
	const active = tabs.some((tab) => tab.id === raw.active) ? (raw.active as string) : null;
	return { tabs, active };
}

export function openSuiteTab(
	state: SuiteTabState,
	app: string,
	framework: string,
	id: string,
): SuiteTabState {
	return { tabs: [...state.tabs, { id, app, framework }], active: id };
}

/** Activate a tab; an unknown id or null means Home. */
export function activateSuiteTab(state: SuiteTabState, id: string | null): SuiteTabState {
	return { tabs: state.tabs, active: id && state.tabs.some((t) => t.id === id) ? id : null };
}

/**
 * Close a tab. Closing the active one activates its right neighbour, else the left one, else
 * Home, as browser tab strips do.
 */
export function closeSuiteTab(state: SuiteTabState, id: string): SuiteTabState {
	const index = state.tabs.findIndex((tab) => tab.id === id);
	if (index < 0) return state;
	const tabs = state.tabs.filter((tab) => tab.id !== id);
	if (state.active !== id) return { tabs, active: state.active };
	const next = tabs[index] ?? tabs[index - 1];
	return { tabs, active: next ? next.id : null };
}

export function setSuiteTabFramework(
	state: SuiteTabState,
	id: string,
	framework: string,
): SuiteTabState {
	return {
		tabs: state.tabs.map((tab) => (tab.id === id ? { ...tab, framework } : tab)),
		active: state.active,
	};
}

/** Titles for the strip: the app name, numbered from the second tab of the same app on. */
export function suiteTabTitles(
	tabs: readonly SuiteTab[],
	nameOf: (app: string) => string,
): Map<string, string> {
	const seen = new Map<string, number>();
	const titles = new Map<string, string>();
	for (const tab of tabs) {
		const count = (seen.get(tab.app) ?? 0) + 1;
		seen.set(tab.app, count);
		titles.set(tab.id, count === 1 ? nameOf(tab.app) : `${nameOf(tab.app)} ${count}`);
	}
	return titles;
}

/** An id no live tab uses. */
export function newSuiteTabId(tabs: readonly SuiteTab[]): string {
	let n = tabs.length + 1;
	while (tabs.some((tab) => tab.id === `t${n}`)) n++;
	return `t${n}`;
}
