/**
 * Host-hidden ribbon tabs (`customization.ribbon.hiddenTabs`) on the shared `office-ui-ribbon`:
 * `file` hides the File button (its key tip goes with it) and every other id hides that native
 * tab's panel for the `host` reason, beside the product's own reasons (contextual, Customize
 * Ribbon, Options), so none of them undoes another. Host add-in tabs are left alone. Word, Excel
 * and Visio share this; each product also keeps its backstage shut while File is hidden.
 */
import { FILE_RIBBON_TAB } from 'ooxml-core/embedding';

/** The reason a tab the host removed carries on its panel (`data-tab-hidden-by`). */
export const HOST_TAB_HIDDEN_REASON = 'host';

type HostTabsRibbon = HTMLElement & {
	noFile: boolean;
	setTabHidden?(tab: string, reason: string, hidden: boolean): void;
};

/** Applies the host's hidden tabs to `ribbon`; an empty set shows every native tab again. */
export function applyHostRibbonTabs(ribbon: HTMLElement, hidden: ReadonlySet<string>): void {
	const element = ribbon as HostTabsRibbon;
	const noFile = hidden.has(FILE_RIBBON_TAB);
	if (element.noFile !== noFile) element.noFile = noFile;
	for (const panel of ribbon.children) {
		if (!(panel instanceof HTMLElement)) continue;
		const tab = panel.dataset.ribbonTab;
		if (!tab || 'addIn' in panel.dataset) continue;
		element.setTabHidden?.(tab, HOST_TAB_HIDDEN_REASON, hidden.has(tab));
	}
}

/**
 * True when `node` sits in a native tab the host removed, so ribbon search (Tell me) can skip a
 * command whose only home is there. A copy of the command in a visible tab still counts.
 */
export function inHostHiddenTab(node: Element): boolean {
	const panel = node.closest<HTMLElement>('[data-ribbon-tab]');
	return (panel?.dataset.tabHiddenBy ?? '').split(' ').includes(HOST_TAB_HIDDEN_REASON);
}
