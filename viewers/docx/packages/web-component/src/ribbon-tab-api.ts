/**
 * The tab row belongs to the shared `office-ui-ribbon`: it builds the tabs from the panels, keeps one
 * panel visible and renders the File button. Everything docx does to a tab is therefore done to its
 * panel, through these helpers, so no other module reaches into the ribbon's shadow root.
 */
export type RibbonElement = HTMLElement & {
	selected: string;
	focusTab(): void;
	focusFile(): void;
	setTabHidden(tab: string, reason: string, hidden: boolean): void;
	tabButton(tab: string): HTMLButtonElement | null;
	fileButton(): HTMLButtonElement | null;
};

/** The panel of tab `key` (`home`, `table`, `header-footer`...). */
export function panelOf(ribbon: ParentNode, key: string): HTMLElement | null {
	return ribbon.querySelector<HTMLElement>(`[data-ribbon-tab="${key}"]`);
}

/** Every panel, in tab order. */
export function panelsOf(ribbon: ParentNode): HTMLElement[] {
	return [...ribbon.querySelectorAll<HTMLElement>(':scope > [data-ribbon-tab]')];
}

/** The key of the tab that is showing. */
export function selectedTab(ribbon: HTMLElement): string {
	return (ribbon as RibbonElement).selected;
}

export function selectTab(ribbon: HTMLElement, key: string): void {
	(ribbon as RibbonElement).selected = key;
}

/** Why a tab can be missing: the tools do not apply to the selection, or a customisation removed it. */
export type TabHiddenReason = 'contextual' | 'custom';

/** Hides or shows tab `key` for one reason; the shared ribbon shows it once no reason hides it. */
export function setTabHidden(
	ribbon: HTMLElement,
	key: string,
	reason: TabHiddenReason,
	hidden: boolean,
): void {
	(ribbon as RibbonElement).setTabHidden(key, reason, hidden);
}

export const isTabHidden = (panel: HTMLElement): boolean => panel.hasAttribute('data-tab-hidden');

/** The rendered tab buttons, for focus and key tips only. */
export function tabButtons(ribbon: HTMLElement): HTMLButtonElement[] {
	return panelsOf(ribbon).flatMap((panel) => {
		const button = (ribbon as RibbonElement).tabButton(panel.dataset.ribbonTab!);
		return button ? [button] : [];
	});
}

/** The File button the ribbon renders. */
export function fileButton(ribbon: HTMLElement): HTMLButtonElement | null {
	return (ribbon as RibbonElement).fileButton();
}
