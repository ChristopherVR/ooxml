/**
 * The PowerPoint ribbon tab row as one descriptor for every binding: which tabs show, in which
 * order, with which label, tooltip and kind, and which one is selected. The shared
 * `pptx-ui-ribbon-tabs` element (the shared `office-ui-ribbon` tab row Word, Excel and Visio use)
 * draws it, so a binding only builds this state from its own stores and routes the element's
 * `tab-select` and `collapse-toggle` events.
 *
 * The order is PowerPoint's: the fixed tabs (File first), the host's add-in tabs after them, then
 * the contextual tabs the selection brings up (Shape Format, Picture Format, Table Design, ...).
 *
 * @module render/ribbon-tab-row
 */

import type { RibbonAddInTab } from './ribbon-add-ins';
import { visibleRibbonAddIns } from './ribbon-add-ins';
import { contextualTabLabelKey } from './ribbon-galleries/gallery-placements';
import type {
	RibbonContextualTabId,
	ToolbarActionId,
	ToolbarTabDefinition,
} from './toolbar-actions';
import { filterVisibleTabs, TOOLBAR_TABS } from './toolbar-actions';

/** What a tab of the row stands for; it tells the binding which of its stores to update. */
export type RibbonTabRowKind = 'file' | 'tab' | 'addIn' | 'contextual';

export interface RibbonTabRowItem {
	readonly id: string;
	readonly label: string;
	readonly kind: RibbonTabRowKind;
	/** Drawn in the contextual tint. */
	readonly contextual?: boolean;
	readonly activateOnFocus?: false;
	/** The ScreenTip; absent while Options > General turns ScreenTips off. */
	readonly title?: string;
}

export interface RibbonTabRowState {
	readonly items: readonly RibbonTabRowItem[];
	readonly manifestLocale?: string;
	/** The selected tab's id, or empty when the active tab is not in the row. */
	readonly selected: string;
	/** The tab list's accessible name. */
	readonly label: string;
}

export interface RibbonTabRowInput {
	readonly translate: (key: string) => string;
	/** The tab shown: a fixed, add-in or contextual tab id (`file` for the backstage). */
	readonly active: string | null | undefined;
	/** The host's `hiddenActions`: a listed tab id leaves its tab out (File included). */
	readonly hiddenActions?: readonly ToolbarActionId[] | undefined;
	/** Options > Customize Ribbon: tabs the user unticked. File always stays. */
	readonly hiddenTabIds?: Iterable<string> | undefined;
	/** The host's tabs (`ribbonAddIns`); unusable ids are dropped. */
	readonly addInTabs?: readonly RibbonAddInTab[] | null | undefined;
	/** The contextual tabs the selection brings up, in PowerPoint's order. */
	readonly contextualTabs?: readonly RibbonContextualTabId[] | undefined;
	/** The ScreenTip for a tab label (`resolveScreenTip`); undefined draws no tooltip. */
	readonly screenTip?: ((label: string) => string | undefined) | undefined;
	/** The fixed tabs; PowerPoint's by default. */
	readonly tabs?: readonly ToolbarTabDefinition[] | undefined;
}

/** The translation key of the tab list's accessible name. */
export const RIBBON_TAB_ROW_LABEL_KEY = 'pptx.ribbon.tabsAria';

/** Build the tab row from the binding's state. Pure: the same input gives an equal state. */
export function buildRibbonTabRow(input: RibbonTabRowInput): RibbonTabRowState {
	const screenTip = input.screenTip ?? ((label: string) => label);
	const userHidden = new Set(input.hiddenTabIds ?? []);
	userHidden.delete('file');
	const item = (id: string, label: string, kind: RibbonTabRowKind): RibbonTabRowItem => {
		const title = screenTip(label);
		return {
			id,
			label,
			kind,
			...(kind === 'contextual' ? { contextual: true } : {}),
			...(kind === 'file' ? { activateOnFocus: false } : {}),
			...(title ? { title } : {}),
		};
	};
	const items: RibbonTabRowItem[] = [
		...filterVisibleTabs(input.tabs ?? TOOLBAR_TABS, input.hiddenActions)
			.filter((tab) => !userHidden.has(tab.id))
			.map((tab) =>
				item(tab.id, input.translate(tab.labelKey), tab.id === 'file' ? 'file' : 'tab'),
			),
		// The host's own label: it is not a translation key.
		...visibleRibbonAddIns(input.addInTabs).map((tab) => item(tab.id, tab.label, 'addIn')),
		...(input.contextualTabs ?? []).map((id) =>
			item(id, input.translate(contextualTabLabelKey(id)), 'contextual'),
		),
	];
	const active = input.active ?? '';
	return {
		items,
		selected: items.some((tab) => tab.id === active) ? active : '',
		label: input.translate(RIBBON_TAB_ROW_LABEL_KEY),
		manifestLocale: input.translate('pptx.addIns.manifestLocale'),
	};
}

/** The kind of tab `id` in `state`, for routing a selection; undefined for an unknown id. */
export function ribbonTabRowKind(
	state: RibbonTabRowState,
	id: string,
): RibbonTabRowKind | undefined {
	return state.items.find((tab) => tab.id === id)?.kind;
}

/** Detail of the `tab-select` event `pptx-ui-ribbon-tabs` fires. */
export interface RibbonTabSelectDetail {
	readonly tab: string;
	readonly kind: RibbonTabRowKind;
}
