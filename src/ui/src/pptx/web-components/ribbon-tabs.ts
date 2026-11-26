import { defineElement } from '../../registry';
import { defineRibbon, type OfficeUiRibbon } from '../../ribbon/ribbon-tabs';
import type { RibbonAddInTab } from '../../ribbon/add-in-tabs';
import type { RibbonTabRowState, RibbonTabSelectDetail } from '../render';
import { RIBBON_ADD_IN_TAB_ATTR, RIBBON_CONTEXTUAL_TAB_ATTR } from '../render';
import { attachControlStyles } from './control-styles';
import { bridgeCss } from './office-token-bridge';
import { RIBBON_TABS_CSS } from './ribbon-tabs-styles';
import { trackSmartArtCanvasSelection } from './smartart-canvas-selection';
import { preserveInlineSymbolCaret } from '../render/inline-symbol';

/** The `pptx-ui-ribbon-tabs` element as a binding sees it. */
export interface PptxUiRibbonTabsElement extends HTMLElement {
	state: RibbonTabRowState | undefined;
	/** The tabs as drawn. */
	tabButtons(): HTMLButtonElement[];
	setInstalledAddInTabs(tabs: readonly RibbonAddInTab[]): void;
}

/**
 * The tab buttons of every `pptx-ui-ribbon-tabs` under `root` (they live in its shadow root), for
 * tests and for code that looks a tab up by its label.
 */
export function ribbonTabButtons(root: ParentNode): HTMLButtonElement[] {
	const hosts = [...root.querySelectorAll<HTMLElement>('pptx-ui-ribbon-tabs')];
	if (root instanceof HTMLElement && root.localName === 'pptx-ui-ribbon-tabs') hosts.unshift(root);
	return hosts.flatMap((host) => [
		...(host.shadowRoot?.querySelectorAll<HTMLButtonElement>('[role="tab"]') ?? []),
	]);
}

/** The first tab under `root` that matches `selector` (any tab by default), or null. */
export function ribbonTab(root: ParentNode, selector = '[role="tab"]'): HTMLButtonElement | null {
	return ribbonTabButtons(root).find((tab) => tab.matches(selector)) ?? null;
}

/**
 * `pptx-ui-ribbon-tabs`: the PowerPoint ribbon tab row, the shared `office-ui-ribbon` (the tab row
 * Word, Excel and Visio draw) under the pptx tag, with the pptx token bridge. Controlled through
 * `state` (`buildRibbonTabRow`): the element never changes the selection itself. Events:
 * `tab-select` `{ tab, kind }` (a click, or the arrow keys, Home and End on the tabs) and
 * `collapse-toggle` (a double click on a tab other than File). Light-DOM children with
 * `slot="actions"` sit at the right end of the row (Record, Comments and Share, Ribbon Display
 * Options).
 */
export function definePptxRibbonTabs(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-ribbon-tabs')) {
		return;
	}
	defineRibbon(registry);
	const Base = registry.get('office-ui-ribbon') as unknown as typeof OfficeUiRibbon;
	class PptxRibbonTabs extends Base {
		#model: RibbonTabRowState | undefined;
		#canvasSelection: AbortController | undefined;
		#manifestTabs: readonly RibbonAddInTab[] = [];
		#manifestTab: string | undefined;
		#syncManifestPanels: typeof import('../../ribbon/add-in-tabs').syncRibbonAddIns | undefined;
		constructor() {
			super();
			this.addEventListener('office-ribbon-select', this.#onSelect);
			if (this.shadowRoot) {
				this.shadowRoot.addEventListener('pointerdown', (event) => {
					if (event.target instanceof HTMLElement && event.target.closest('[data-tab="insert"]')) {
						preserveInlineSymbolCaret(event, this.ownerDocument);
					}
				});
				// On the shadow root, so a double click that is not composed still counts.
				this.shadowRoot.addEventListener('dblclick', this.#onDoubleClick);
				attachControlStyles(this.shadowRoot, bridgeCss('pptx-ui-ribbon-tabs') + RIBBON_TABS_CSS);
			}
		}
		override connectedCallback(): void {
			super.connectedCallback();
			this.#canvasSelection = new AbortController();
			trackSmartArtCanvasSelection(this, this.#canvasSelection.signal);
		}
		override disconnectedCallback(): void {
			this.#canvasSelection?.abort();
			super.disconnectedCallback();
		}
		get state(): RibbonTabRowState | undefined {
			return this.#model;
		}
		set state(value: RibbonTabRowState | undefined) {
			this.#model = value;
			this.label = value?.label ?? null;
			this.#applyState();
		}
		setInstalledAddInTabs(tabs: readonly RibbonAddInTab[]): void {
			this.#manifestTabs = tabs;
			if (this.#syncManifestPanels || !tabs.length) this.#applyState();
			else
				void import('../../ribbon/add-in-tabs').then(({ syncRibbonAddIns }) => {
					if (!this.isConnected) return;
					this.#syncManifestPanels = syncRibbonAddIns;
					this.#applyState();
				});
		}
		#applyState(): void {
			const items = this.#model?.items ?? [];
			if (!this.#syncManifestPanels) {
				this.items = items;
				this.selected = this.#model?.selected ?? '';
				return;
			}
			const ids = new Set(items.map(({ id }) => id));
			const local = this.#manifestTabs.filter(({ id }) => !ids.has(id));
			this.#syncManifestPanels(this, local);
			if (!local.some(({ id }) => id === this.#manifestTab)) this.#manifestTab = undefined;
			this.items = [...items, ...local];
			this.selected = this.#manifestTab ?? this.#model?.selected ?? '';
			for (const panel of this.children) {
				if (!(panel instanceof HTMLElement) || !panel.hasAttribute('data-add-in')) continue;
				panel.hidden = panel.dataset.ribbonTab !== this.#manifestTab;
			}
		}
		/** The tabs as drawn, for tests and focus handling. */
		tabButtons(): HTMLButtonElement[] {
			return [...this.renderRoot.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
		}
		#onSelect = (event: Event): void => {
			// The binding owns the selection: it answers with a new `state`.
			event.preventDefault();
			const tab = (event as CustomEvent<{ tab: string }>).detail.tab;
			const kind = this.#model?.items.find((item) => item.id === tab)?.kind;
			if (!kind && this.#manifestTabs.some(({ id }) => id === tab)) {
				this.#manifestTab = tab;
				this.#applyState();
				return;
			}
			if (kind) {
				if (this.#manifestTab) {
					this.#manifestTab = undefined;
					this.#applyState();
				}
				this.dispatchEvent(
					new CustomEvent<RibbonTabSelectDetail>('tab-select', {
						detail: { tab, kind },
						bubbles: true,
						composed: true,
					}),
				);
			}
		};
		#onDoubleClick = (event: Event): void => {
			const tab = event
				.composedPath()
				.find(
					(node): node is HTMLElement =>
						node instanceof HTMLElement && node.getAttribute('role') === 'tab',
				);
			if (tab && tab.dataset.tab !== 'file') {
				this.dispatchEvent(new CustomEvent('collapse-toggle', { bubbles: true, composed: true }));
			}
		};
		protected override updated(): void {
			super.updated();
			// The markers the bindings' tab rows carried, so selectors and tests keep working.
			const kinds = new Map(this.#model?.items.map((item) => [item.id, item.kind]));
			for (const button of this.tabButtons()) {
				const id = button.dataset.tab ?? '';
				const kind = kinds.get(id);
				button.toggleAttribute(RIBBON_ADD_IN_TAB_ATTR, kind === 'addIn');
				button.toggleAttribute(RIBBON_CONTEXTUAL_TAB_ATTR, kind === 'contextual');
				if (kind === 'addIn') button.setAttribute(RIBBON_ADD_IN_TAB_ATTR, id);
				if (kind === 'contextual') button.setAttribute(RIBBON_CONTEXTUAL_TAB_ATTR, id);
			}
		}
	}
	defineElement(
		registry,
		'pptx-ui-ribbon-tabs',
		PptxRibbonTabs as unknown as CustomElementConstructor,
	);
}
