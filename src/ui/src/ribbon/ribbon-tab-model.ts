/**
 * One tab a host gives `items` directly: the tab row then draws these instead of reading tabs off
 * the panels (a product that renders its panels elsewhere, as PowerPoint's bindings do).
 */
export interface OfficeRibbonTabItem {
	readonly id: string;
	readonly label: string;
	/** Tints the tab as a contextual tab (`--office-ribbon-contextual`). */
	readonly contextual?: boolean;
	readonly keytip?: string;
	/** A tab that opens a dialog can require explicit activation after moving focus. */
	readonly activateOnFocus?: false;
	/** The tab's tooltip (a ScreenTip); none when absent. */
	readonly title?: string;
}

export interface RibbonTab {
	id: string;
	label: string;
	contextual: boolean;
	keytip?: string;
	keytipPanel?: string;
	title?: string;
}

/** The tabs of a host-given `items` list, in its order. */
export function tabsFromItems(items: readonly OfficeRibbonTabItem[]): RibbonTab[] {
	return items.map((item) => ({
		id: item.id,
		label: item.label,
		contextual: item.contextual === true,
		...(item.keytip ? { keytip: item.keytip } : {}),
		...(item.title ? { title: item.title } : {}),
	}));
}

/** The tabs read off the shown panels, marking each panel as the tab panel it is. */
export function tabsFromPanels(panels: readonly HTMLElement[]): RibbonTab[] {
	return panels.map((panel) => {
		const id = panel.dataset.ribbonTab!;
		panel.setAttribute('role', 'tabpanel');
		panel.setAttribute('aria-label', panel.dataset.label ?? id);
		return {
			id,
			label: panel.dataset.label ?? id,
			contextual: panel.dataset.contextual !== undefined,
			...(panel.dataset.tabKeytip
				? {
						keytip: panel.dataset.tabKeytip,
						...(panel.id ? { keytipPanel: panel.id } : {}),
					}
				: {}),
		};
	});
}

/** The tab a navigation key moves to from `at`: Home, End and the arrows, wrapping. */
export function tabForKey<T>(tabs: readonly T[], at: number, key: string): T | undefined {
	if (key === 'Home') return tabs[0];
	if (key === 'End') return tabs.at(-1);
	const step = { ArrowRight: 1, ArrowLeft: -1 }[key];
	return step ? tabs[(at + step + tabs.length) % tabs.length] : undefined;
}

/** Scrolls the tab strip just enough to show `active` whole. */
export function revealTab(active: HTMLElement | null): void {
	const list = active?.parentElement;
	if (!active || !list || list.scrollWidth <= list.clientWidth) return;
	const tab = active.getBoundingClientRect();
	const view = list.getBoundingClientRect();
	if (tab.left < view.left) list.scrollLeft -= view.left - tab.left;
	else if (tab.right > view.right) list.scrollLeft += tab.right - view.right;
}
