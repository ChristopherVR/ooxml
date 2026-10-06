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
	/** Pinned tabs sit first, show only their icon and cannot be closed by accident. */
	readonly pinned: boolean;
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
		const { id, app, framework, pinned } = item as Record<string, unknown>;
		if (typeof id !== 'string' || typeof app !== 'string' || typeof framework !== 'string')
			continue;
		if (tabs.some((tab) => tab.id === id)) continue;
		tabs.push({ id, app, framework, pinned: pinned === true });
	}
	const active = tabs.some((tab) => tab.id === raw.active) ? (raw.active as string) : null;
	return { tabs: pinnedFirst(tabs), active };
}

export function openSuiteTab(
	state: SuiteTabState,
	app: string,
	framework: string,
	id: string,
): SuiteTabState {
	return { tabs: [...state.tabs, { id, app, framework, pinned: false }], active: id };
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

/** Stable partition: pinned tabs first, every group keeping its order. */
export function pinnedFirst(tabs: readonly SuiteTab[]): SuiteTab[] {
	return [...tabs.filter((t) => t.pinned), ...tabs.filter((t) => !t.pinned)];
}

/** Pin a tab (it joins the end of the pinned group) or unpin it (it leads the unpinned ones). */
export function pinSuiteTab(state: SuiteTabState, id: string, pinned: boolean): SuiteTabState {
	const tab = state.tabs.find((t) => t.id === id);
	if (!tab || tab.pinned === pinned) return state;
	const rest = state.tabs.filter((t) => t.id !== id);
	const moved = { ...tab, pinned };
	const pinnedCount = rest.filter((t) => t.pinned).length;
	rest.splice(pinned ? pinnedCount : pinnedCount, 0, moved);
	return { tabs: rest, active: state.active };
}

/**
 * Move a tab to just before `beforeId` (the end when null). A tab never leaves its group: a
 * pinned tab stays among the pinned ones and an unpinned tab among the others.
 */
export function moveSuiteTab(
	state: SuiteTabState,
	id: string,
	beforeId: string | null,
): SuiteTabState {
	const tab = state.tabs.find((t) => t.id === id);
	if (!tab || beforeId === id) return state;
	const rest = state.tabs.filter((t) => t.id !== id);
	const pinnedCount = rest.filter((t) => t.pinned).length;
	const wanted = beforeId === null ? rest.length : rest.findIndex((t) => t.id === beforeId);
	if (wanted < 0) return state;
	const index = tab.pinned ? Math.min(wanted, pinnedCount) : Math.max(wanted, pinnedCount);
	rest.splice(index, 0, tab);
	return { tabs: rest, active: state.active };
}

function without(
	state: SuiteTabState,
	keep: (tab: SuiteTab, index: number) => boolean,
	anchor: string,
) {
	const tabs = state.tabs.filter(keep);
	return { tabs, active: tabs.some((t) => t.id === state.active) ? state.active : anchor };
}

/** Close every tab but this one; pinned tabs stay. */
export function closeOtherSuiteTabs(state: SuiteTabState, id: string): SuiteTabState {
	return without(state, (t) => t.id === id || t.pinned, id);
}

/** Close the tabs after this one; pinned tabs stay. */
export function closeSuiteTabsToRight(state: SuiteTabState, id: string): SuiteTabState {
	const at = state.tabs.findIndex((t) => t.id === id);
	if (at < 0) return state;
	return without(state, (t, i) => i <= at || t.pinned, id);
}
