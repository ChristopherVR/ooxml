import { ribbonTabCueLabel } from '../../ribbon/ribbon-tab-cue-label';
/**
 * The ribbon renderer: the registered tabs (contextual ones only while their test passes), one panel per
 * tab with labelled groups and dialog launchers, KeyTips, hidden actions and overflow folding. The tab
 * row, File button, collapse and peek belong to the shared `office-ui-ribbon`. Structure and look follow docx-viewer's ribbon.ts.
 */
import { syncRibbonAddIns, type RibbonAddInTab } from '../../ribbon/add-in-tabs';
import { applyHostRibbonTabs } from '../../embedding/host-ribbon-tabs';
import { normalizeEditorLocale, type EditorContext } from 'ooxml-core/xlsx/ui';
import { defineRibbon } from '../../controls';
import { el, type RenderScope } from './controls';
import { renderGroup } from './render-group';
import { attachRibbonOverflow, closeRibbonOverflow, refitRibbon } from './overflow';
import { onRibbonTabsChange, ribbonTabs, type RibbonTab } from 'ooxml-core/xlsx/ui';
import { closeRibbonPopover } from './popover';
import { showTabKeyTips } from './keytips';
import { isTabHidden, selectTab, setTabHidden, type RibbonElement } from 'ooxml-core/xlsx/ui';
import { applyRibbonVisibility } from './visibility';
import { wireCellStyleGalleries } from './cell-style-gallery';

export interface RibbonHandlers {
	openBackstage(): void;
	isHidden(id: string): boolean;
	/** The native tabs the host removed (`customization.ribbon.hiddenTabs`), `file` included. */
	hiddenTabs?(): ReadonlySet<string>;
}

export interface Ribbon {
	readonly element: HTMLElement;
	/** Re-reads every control state (enabled, pressed, values) and the contextual tabs. */
	refresh(): void;
	/** Re-applies the tabs the host removed (`customization.ribbon.hiddenTabs`). */
	applyHiddenTabs(): void;
	/** Rebuilds all tabs (after a locale change or newly registered tabs). */
	rebuild(): void;
	selectTab(id: string): void;
	/** Tabs a host adds after Excel's own (`ribbonAddIns` of the editor). */
	setAddIns(tabs: readonly RibbonAddInTab[]): void;
	activeTab(): string | undefined;
	/** Alt / F10: focus the selected tab and show the tab KeyTips. */
	showKeyTips(): void;
	focus(): void;
	/** Follows tab registrations and panel sizes; tabs registered while away are built now. */
	connect(): void;
	/** Lets go of everything outside the ribbon, so a removed editor can be collected. */
	disconnect(): void;
}

export function createRibbon(ctx: EditorContext, handlers: RibbonHandlers): Ribbon {
	defineRibbon();
	const doc = ctx.host.ownerDocument;
	const root = doc.createElement('office-ui-ribbon') as RibbonElement;
	root.addEventListener('pointerdown', (event) => {
		if (
			ctx.grid()?.isEditing() &&
			event
				.composedPath()
				.some((node) => node instanceof Element && node.matches('[role="tab"][data-tab="insert"]'))
		)
			event.preventDefault();
	});
	root.className = 'xve-ribbon';
	root.setAttribute('part', 'ribbon');
	root.setAttribute('role', 'toolbar');
	root.setAttribute('file-expanded', 'false');
	root.toggleAttribute('collapsible', true);
	let scope: RenderScope = { ctx, doc, updates: [], isHidden: handlers.isHidden };
	const deferred = new Map<HTMLElement, RenderScope>();
	let tabs: RibbonTab[] = [];
	let stopListening: (() => void) | undefined;
	let stopOverflow = () => {};
	let stopGalleries = () => {};
	let addIns: readonly RibbonAddInTab[] = [];
	const syncAddIns = () => syncRibbonAddIns(root, addIns, { panelClass: 'ribbon-panel' });

	const panels = () => [
		...[...root.children].filter((node): node is HTMLElement =>
			node.classList.contains('ribbon-panel'),
		),
	];
	const panelFor = (id: string) => panels().find((panel) => panel.dataset.tab === id);

	const select = (id: string) => {
		const panel = panelFor(id);
		if (panel && !isTabHidden(panel)) selectTab(root, id);
	};

	// The shared ribbon announces a tab change before it applies it: fold the new panel once it shows.
	const refit = () => void root.updateComplete.then(() => refitRibbon(root));
	root.addEventListener('office-ribbon-select', () => {
		closeRibbonPopover();
		refit();
	});
	root.addEventListener('office-ribbon-collapse', refit);
	root.addEventListener('office-ribbon-file', () => handlers.openBackstage());

	const renderGroups = (panel: HTMLElement, tab: RibbonTab, tabScope: RenderScope) => {
		for (const group of tab.groups) {
			const node = renderGroup(tabScope, group);
			if (node) panel.append(node);
		}
	};

	const build = () => {
		unobserve();
		closeRibbonPopover();
		scope = { ctx, doc, updates: [], isHidden: handlers.isHidden };
		deferred.clear();
		tabs = [...ribbonTabs()];
		// The overflow chevron reads its label from here, so a locale change reaches it on rebuild.
		root.dataset.editorLocale = normalizeEditorLocale(ctx.locale());
		root.setAttribute('aria-label', ctx.t('Ribbon'));
		root.setAttribute('label', ctx.t('Ribbon tabs'));
		root.setAttribute('file-label', ctx.t('File'));
		root.setAttribute('overflow-label', ribbonTabCueLabel(ctx.locale()));
		root.setAttribute('collapse-label', ctx.t('Collapse the ribbon'));
		for (const panel of panels()) panel.remove();
		for (const tab of tabs) {
			const panel = el(doc, 'div', 'ribbon-panel');
			panel.id = `xve-panel-${tab.id}`;
			panel.dataset.tab = tab.id;
			panel.dataset.ribbonTab = tab.id;
			panel.dataset.label = ctx.t(tab.label);
			if (tab.contextual) panel.dataset.contextual = '';
			if (!tab.deferControls) renderGroups(panel, tab, scope);
			root.append(panel);
		}
		syncAddIns();
		refresh();
		if (stopListening) observe();
	};

	const observe = () => {
		stopOverflow = attachRibbonOverflow(root);
		stopGalleries = wireCellStyleGalleries(root);
	};

	const unobserve = () => {
		stopOverflow();
		stopGalleries();
		stopOverflow = stopGalleries = () => {};
	};

	const applyHiddenTabs = () => applyHostRibbonTabs(root, handlers.hiddenTabs?.() ?? new Set());

	const refresh = () => {
		for (const update of scope.updates) update();
		for (const tab of tabs) {
			let show = true;
			try {
				show = tab.contextual ? tab.contextual(ctx) : true;
			} catch {
				show = false;
			}
			if (show && tab.deferControls) {
				const panel = panelFor(tab.id);
				if (panel) {
					let tabScope = deferred.get(panel);
					if (!tabScope) {
						tabScope = { ...scope, updates: [] };
						renderGroups(panel, tab, tabScope);
						deferred.set(panel, tabScope);
					}
					for (const update of tabScope.updates) update();
				}
			}
			setTabHidden(root, tab.id, 'contextual', !show);
		}
		applyHiddenTabs();
		applyRibbonVisibility(root, handlers.isHidden);
	};

	build();
	return {
		element: root,
		refresh,
		applyHiddenTabs,
		rebuild: build,
		selectTab: select,
		setAddIns: (next) => {
			addIns = next;
			const before = root.querySelector('[data-add-in]');
			syncAddIns();
			// Equal tabs keep their panels; only new panels need the overflow measured again.
			if (root.querySelector('[data-add-in]') === before || !stopListening) return;
			stopOverflow();
			stopOverflow = attachRibbonOverflow(root);
		},
		activeTab: () => root.selected || undefined,
		showKeyTips: () => {
			root.focusTab();
			showTabKeyTips(root, select);
		},
		focus: () => root.focusTab(),
		connect: () => {
			if (stopListening) return;
			stopListening = onRibbonTabsChange(build);
			const current = ribbonTabs();
			if (current.length === tabs.length && current.every((tab, index) => tab === tabs[index]))
				observe();
			else build();
		},
		disconnect: () => {
			stopListening?.();
			stopListening = undefined;
			unobserve();
			closeRibbonPopover(ctx.host);
			closeRibbonOverflow(ctx.host);
		},
	};
}
