import type { VisioShape } from 'ooxml-core/visio';
import type { ViewerState } from './controller';
import { selectedShape } from './shape-inspector';
import { orgChartSelection } from './viewer-org-chart';
import { swimlaneSelection } from './viewer-swimlanes';

/** The format tabs, in the order the ribbon lists them; one at most shows at a time. */
export const CONTEXTUAL_TABS = [
	'shape-format',
	'picture-format',
	'container-format',
	'chart-design',
] as const;
export type ContextualTab = (typeof CONTEXTUAL_TABS)[number];
export type SelectionKind = 'picture' | 'container' | 'chart' | 'shape';

const MIRRORED = ['disabled', 'main-disabled', 'checked', 'pressed', 'title'] as const;

/** The kind of object that decides which specialised tab a single selection gets. */
export function selectionKind(shape: VisioShape | undefined): SelectionKind {
	if (shape?.embeddedObject) return 'chart';
	if (shape?.image) return 'picture';
	if (shape?.structure?.type === 'container') return 'container';
	return 'shape';
}

/**
 * The one contextual tab for the selection. A lone picture, container or chart gets its own tab
 * (which carries the Arrange and Size controls too) instead of Shape Format; every other
 * selection, including several shapes, gets Shape Format; no selection gets none.
 */
export function contextualTabFor(state: ViewerState): ContextualTab | undefined {
	if (state.selectedShapes.length === 0) return undefined;
	if (state.selectedShapes.length === 1) {
		const kind = selectionKind(selectedShape(state.document, state.selectedShape, state.pageIndex));
		if (kind === 'picture') return 'picture-format';
		if (kind === 'container') return 'container-format';
		if (kind === 'chart') return 'chart-design';
	}
	return 'shape-format';
}

/**
 * The state of the contextual tabs. Their Shape Styles, Arrange and Size controls are the Home
 * controls built again: every module that syncs a control by its id updates the first match
 * (Home), so this copies that state onto the contextual twins instead of teaching each module
 * about several copies. Controls that exist only on a contextual tab are synced by their owner.
 */
export function renderContextualTabs(toolbar: HTMLElement, state: ViewerState): void {
	const ribbon = toolbar as HTMLElement & { setTabHidden?(t: string, r: string, h: boolean): void };
	const active = contextualTabFor(state);
	for (const tab of CONTEXTUAL_TABS) ribbon.setTabHidden?.(tab, 'contextual', tab !== active);
	// The Cross-Functional Flowchart tab shows beside the format tab while a flowchart, one of
	// its lanes or a shape in a lane is selected; its commands are its own, so nothing is mirrored.
	ribbon.setTabHidden?.('cross-functional', 'contextual', swimlaneSelection(state) === undefined);
	// The Org Chart tab shows beside the format tab while a position is selected.
	ribbon.setTabHidden?.('org-chart', 'contextual', orgChartSelection(state) === undefined);
	const panel = active && toolbar.querySelector<HTMLElement>(`[data-ribbon-tab="${active}"]`);
	if (!panel) return;
	const selector = '[command], [data-menu]';
	const twins = [...panel.querySelectorAll<HTMLElement>(selector)].map((control) => ({
		control,
		id: control.getAttribute('command') ?? control.getAttribute('data-menu'),
	}));
	const ids = new Set(twins.map(({ id }) => id));
	const sources = new Map<string, Map<string, HTMLElement>>();
	for (const control of toolbar.querySelectorAll<HTMLElement>(selector)) {
		for (const id of [control.getAttribute('command'), control.getAttribute('data-menu')]) {
			if (id === null || !ids.has(id) || control.closest('[data-contextual]')) continue;
			const byTag = sources.get(id) ?? new Map<string, HTMLElement>();
			if (!byTag.has(control.localName)) byTag.set(control.localName, control);
			sources.set(id, byTag);
		}
	}
	for (const { control: twin, id } of twins) {
		const source = id === null ? undefined : sources.get(id)?.get(twin.localName);
		if (!source) continue;
		for (const name of MIRRORED) {
			const value = source.getAttribute(name);
			if (value === null) twin.removeAttribute(name);
			else if (twin.getAttribute(name) !== value) twin.setAttribute(name, value);
		}
		if (source.localName === 'office-ui-gallery')
			(twin as HTMLElement & { state?: unknown }).state = (
				source as HTMLElement & { state?: unknown }
			).state;
	}
}
