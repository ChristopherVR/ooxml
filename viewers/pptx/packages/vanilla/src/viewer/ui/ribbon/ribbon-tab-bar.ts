import type {
	RibbonAddInTab,
	RibbonContextualTabId,
	RibbonTabRowState,
	RibbonTabSelectDetail,
	ToolbarActionId,
} from 'ooxml-ui/pptx';
import { buildRibbonTabRow, buildTabRowActionsState, isActionHidden } from 'ooxml-ui/pptx';
import { registerPptxWebControls } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import type { RibbonTabId } from './ribbon-types';

export interface RibbonTabBar {
	el: HTMLElement;
	setActive(tab: RibbonTabId | RibbonContextualTabId | string): void;
	/** Show the host's add-in tabs after the fixed tabs and before the contextual ones. */
	setAddInTabs(tabs: readonly RibbonAddInTab[]): void;
	/**
	 * Show the selection's contextual tabs (Shape Format, Picture Format, ...)
	 * after the fixed tabs, PowerPoint-style.
	 */
	setContextualTabs(tabs: readonly RibbonContextualTabId[]): void;
	/** Hide/show tab buttons per Options > Customize Ribbon (File always shown). */
	setHiddenTabs(hidden: ReadonlySet<string>): void;
	/** Set each tab's `title` from a ScreenTip resolver (Options > General). */
	applyScreenTips(resolve: (label: string) => string | undefined): void;
}

/** Right-side quick actions on the tab row (React's `TabRowActions`). */
export interface RibbonTabBarActions {
	/** The red-dot Record button (starts rehearsal/recording). */
	startRecording(): void;
	/** Comments on the shared tab-row element; absent leaves Comments out. */
	openComments?(): void;
}

type TabRowElement = HTMLElement & { state: RibbonTabRowState };

/**
 * The ribbon's tab row: the shared `pptx-ui-ribbon-tabs` (the tab row Word, Excel and Visio
 * draw), fed from `buildRibbonTabRow` whenever a tab list, the active tab or a ScreenTip changes.
 * Its `actions` slot carries the Record button, then the shared `pptx-ui-ribbon-actions`
 * (Comments and Share). This module owns its Comments; the collaboration UI turns on and drives
 * its Share (see `collab/collab-ui.ts`), which stays hidden until then.
 */
export function createRibbonTabBar(
	doc: Document,
	t: Translator,
	onSelect: (tab: RibbonTabId | RibbonContextualTabId | string) => void,
	hiddenActions?: readonly ToolbarActionId[],
	actions?: RibbonTabBarActions,
): RibbonTabBar {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-ribbon-tabs') as TabRowElement;
	el.dataset.pptxChrome = 'ribbon-tabs';
	el.addEventListener('tab-select', (event) =>
		onSelect((event as CustomEvent<RibbonTabSelectDetail>).detail.tab),
	);

	let active: string | null = null;
	let addInTabs: readonly RibbonAddInTab[] = [];
	let contextualTabs: readonly RibbonContextualTabId[] = [];
	let hiddenTabIds: ReadonlySet<string> = new Set();
	let screenTip: ((label: string) => string | undefined) | undefined;
	const render = (): void => {
		el.state = buildRibbonTabRow({
			translate: t,
			active,
			hiddenActions,
			hiddenTabIds,
			addInTabs,
			contextualTabs,
			screenTip,
		});
	};
	render();

	if (actions) {
		const actionsHost = createEl(doc, 'div', 'pptxv-tabrow-actions');
		actionsHost.slot = 'actions';
		if (!isActionHidden('record', hiddenActions)) {
			const record = createEl(doc, 'button', 'pptxv-tabrow-record');
			record.type = 'button';
			record.title = t('pptx.titleBar.record');
			record.setAttribute('aria-label', t('pptx.titleBar.record'));
			const dot = createEl(doc, 'span', 'pptxv-tabrow-record-dot');
			dot.setAttribute('aria-hidden', 'true');
			const label = createEl(doc, 'span');
			label.textContent = t('pptx.titleBar.record');
			record.append(dot, label);
			record.addEventListener('click', () => actions.startRecording());
			actionsHost.appendChild(record);
		}
		const tabActions = doc.createElement('pptx-ui-ribbon-actions');
		tabActions.dataset.pptxChrome = 'tab-row-actions';
		Object.assign(
			tabActions,
			buildTabRowActionsState({
				translate: t,
				showComments: Boolean(actions.openComments),
				commentsOpen: false,
				showShare: false,
				isCollaborating: false,
			}),
		);
		const openComments = actions.openComments;
		if (openComments) {
			tabActions.addEventListener('comments-toggle', () => openComments());
		}
		actionsHost.appendChild(tabActions);
		el.appendChild(actionsHost);
	}

	return {
		el,
		setActive(tab) {
			active = tab;
			render();
		},
		setAddInTabs(tabs) {
			addInTabs = [...tabs];
			render();
		},
		setContextualTabs(tabs) {
			contextualTabs = [...tabs];
			render();
		},
		setHiddenTabs(hidden) {
			hiddenTabIds = hidden;
			render();
		},
		applyScreenTips(resolve) {
			screenTip = resolve;
			render();
		},
	};
}
