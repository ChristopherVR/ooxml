/**
 * The tab row belongs to the shared `office-ui-ribbon`: it builds the tabs from the panels, keeps one
 * panel visible and renders the File button. Products reach a tab through its panel, with these
 * helpers, so no other module touches the ribbon's shadow root.
 */
export type RibbonElement = HTMLElement & {
	selected: string;
	updateComplete: Promise<boolean>;
	focusTab(): void;
	focusFile(): void;
	setTabHidden(tab: string, reason: string, hidden: boolean): void;
	tabButton(tab: string): HTMLButtonElement | null;
	fileButton(): HTMLButtonElement | null;
};

/** Why a tab can be missing: the tools do not apply to the selection, or hidden actions emptied it. */
export type TabHiddenReason = 'contextual' | 'custom';

export const selectTab = (ribbon: HTMLElement, id: string): void => {
	(ribbon as RibbonElement).selected = id;
};

export const setTabHidden = (
	ribbon: HTMLElement,
	id: string,
	reason: TabHiddenReason,
	hidden: boolean,
): void => (ribbon as RibbonElement).setTabHidden(id, reason, hidden);

export const isTabHidden = (panel: HTMLElement): boolean => panel.hasAttribute('data-tab-hidden');

export const tabButton = (ribbon: HTMLElement, id: string): HTMLButtonElement | null =>
	(ribbon as RibbonElement).tabButton(id);
