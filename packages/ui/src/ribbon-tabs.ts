import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/** Emitted when the user picks a tab; cancelable (`preventDefault()` keeps the current tab). */
export type OfficeRibbonSelectEvent = CustomEvent<{ tab: string }>;

const CSS = `
:host { display: block; min-width: 0; }
.head { display: flex; align-items: stretch; min-width: 0; border-bottom: 1px solid var(--office-border, #d1d5db); }
.file { flex: none; min-height: var(--office-ribbon-tab-height, 35px); margin: 0 4px 0 0; padding: 6px 16px; border: 0;
	border-radius: 0; background: var(--office-accent, #2563eb); color: var(--office-accent-foreground, #fff);
	font: inherit; font-size: 13px; cursor: pointer; }
.file:hover { background: color-mix(in srgb, var(--office-accent, #2563eb) 85%, #000); }
.file[hidden] { display: none; }
[role="tablist"] { display: flex; flex: 0 1 auto; align-items: stretch; min-width: 0; overflow-x: auto; scrollbar-width: none; }
[role="tab"] { position: relative; flex: none; min-height: var(--office-ribbon-tab-height, 35px); padding: 6px 12px; border: 0;
	border-radius: 0; background: transparent; color: var(--office-muted-foreground, #6b7280); font: inherit; font-size: 13px;
	cursor: pointer; }
[role="tab"]:hover { color: var(--office-foreground, #1f2937); background: var(--office-surface, #f3f4f6); }
[role="tab"][aria-selected="true"] { color: var(--office-accent, #2563eb); }
[role="tab"][aria-selected="true"]::after { content: ""; position: absolute; inset: auto 0 0; height: 2.5px;
	background: var(--office-accent, #2563eb); }
.file:focus-visible, [role="tab"]:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: -2px; }
::slotted([slot="search"]) { align-self: center; }
@media (max-width: 760px), (pointer: coarse) {
	.file, [role="tab"] { min-height: 44px; }
	.file { min-width: 52px; }
}
@media (forced-colors: active) {
	.file { border: 1px solid ButtonText; }
	[role="tab"][aria-selected="true"]::after { background: Highlight; }
}
`;

/**
 * `<office-ui-ribbon>`: the Office tab row (Quick Access Toolbar, File, tabs, search) over the
 * tab panels. Panels are light-DOM children carrying `data-ribbon-tab` (id), `data-label` and
 * optionally `data-tab-keytip`; the element builds the tabs from them, keeps exactly one panel
 * visible and moves between tabs with the arrow keys, Home and End. Slots: `quick-access`,
 * `search`, `end`. Attributes: `selected`, `label` (tab list name), `file-label` (default "File"),
 * `no-file`, `file-expanded`, `file-keytip`. Events: `office-ribbon-select` `{ tab }`
 * (cancelable) and `office-ribbon-file` when File is activated.
 */
export const defineRibbon = definer('office-ui-ribbon', () => {
	class OfficeUiRibbon extends HTMLElement {
		static observedAttributes = [
			'selected',
			'label',
			'file-label',
			'no-file',
			'file-expanded',
			'file-keytip',
		];
		readonly #file: HTMLButtonElement;
		readonly #tabs: HTMLElement;
		readonly #panelSlot: HTMLSlotElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const head = doc.createElement('div');
			head.className = 'head';
			head.setAttribute('part', 'head');
			const slot = (name: string) => {
				const element = doc.createElement('slot');
				element.name = name;
				return element;
			};
			this.#file = doc.createElement('button');
			this.#file.type = 'button';
			this.#file.className = 'file';
			this.#file.setAttribute('part', 'file');
			this.#file.setAttribute('aria-haspopup', 'dialog');
			this.#file.setAttribute('aria-expanded', 'false');
			this.#file.addEventListener('click', () => emit(this, 'office-ribbon-file', {}));
			this.#tabs = doc.createElement('div');
			this.#tabs.setAttribute('role', 'tablist');
			this.#tabs.setAttribute('part', 'tabs');
			this.#tabs.addEventListener('click', (event) => {
				const tab = (event.target as Element).closest?.<HTMLElement>('[role="tab"]');
				if (tab) this.#choose(tab.dataset.tab!);
			});
			this.#tabs.addEventListener('keydown', (event) => this.#key(event));
			head.append(slot('quick-access'), this.#file, this.#tabs, slot('search'), slot('end'));
			this.#panelSlot = doc.createElement('slot');
			this.#panelSlot.addEventListener('slotchange', () => this.#build());
			root.append(head, this.#panelSlot);
			this.#syncFile();
		}
		connectedCallback(): void {
			this.#build();
		}
		attributeChangedCallback(name: string): void {
			if (name === 'selected') this.#sync();
			else if (name === 'label')
				this.#tabs.setAttribute('aria-label', this.getAttribute('label') ?? 'Ribbon');
			else this.#syncFile();
		}
		get selected(): string {
			return this.getAttribute('selected') ?? this.#panels()[0]?.dataset.ribbonTab ?? '';
		}
		set selected(tab: string) {
			this.setAttribute('selected', tab);
		}
		/** Focus the File button (for example when the backstage it opened closes). */
		focusFile(): void {
			this.#file.focus();
		}
		/** Focus the selected tab. */
		focusTab(): void {
			this.#tabs.querySelector<HTMLElement>('[aria-selected="true"]')?.focus();
		}
		#panels(): HTMLElement[] {
			return [...this.children].filter(
				(child): child is HTMLElement => child instanceof HTMLElement && !!child.dataset.ribbonTab,
			);
		}
		#build(): void {
			const doc = this.ownerDocument;
			if (!this.#tabs.hasAttribute('aria-label'))
				this.#tabs.setAttribute('aria-label', this.getAttribute('label') ?? 'Ribbon');
			this.#tabs.replaceChildren(
				...this.#panels().map((panel) => {
					const id = panel.dataset.ribbonTab!;
					const tab = doc.createElement('button');
					tab.type = 'button';
					tab.setAttribute('role', 'tab');
					tab.id = `tab-${id}`;
					tab.dataset.tab = id;
					tab.textContent = panel.dataset.label ?? id;
					if (panel.dataset.tabKeytip) {
						tab.dataset.keytip = panel.dataset.tabKeytip;
						if (panel.id) tab.dataset.keytipPanel = panel.id;
					}
					panel.setAttribute('role', 'tabpanel');
					panel.setAttribute('aria-label', panel.dataset.label ?? id);
					return tab;
				}),
			);
			this.#sync();
		}
		#sync(): void {
			const selected = this.selected;
			for (const tab of this.#tabs.querySelectorAll<HTMLButtonElement>('[role="tab"]')) {
				const on = tab.dataset.tab === selected;
				tab.setAttribute('aria-selected', String(on));
				tab.tabIndex = on ? 0 : -1;
			}
			for (const panel of this.#panels()) panel.hidden = panel.dataset.ribbonTab !== selected;
		}
		#syncFile(): void {
			this.#file.textContent = this.getAttribute('file-label') ?? 'File';
			this.#file.hidden = this.hasAttribute('no-file');
			this.#file.setAttribute(
				'aria-expanded',
				String(this.getAttribute('file-expanded') === 'true'),
			);
			const keytip = this.getAttribute('file-keytip');
			if (keytip) this.#file.dataset.keytip = keytip;
			else delete this.#file.dataset.keytip;
		}
		#choose(tab: string): void {
			if (tab === this.selected) return;
			if (emit(this, 'office-ribbon-select', { tab }, true)) this.selected = tab;
		}
		#key(event: KeyboardEvent): void {
			const tabs = [...this.#tabs.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
			const at = tabs.findIndex((tab) => tab.dataset.tab === this.selected);
			const step = { ArrowRight: 1, ArrowLeft: -1 }[event.key];
			const next =
				event.key === 'Home'
					? tabs[0]
					: event.key === 'End'
						? tabs.at(-1)
						: step
							? tabs[(at + step + tabs.length) % tabs.length]
							: undefined;
			if (!next) return;
			event.preventDefault();
			this.#choose(next.dataset.tab!);
			this.focusTab();
		}
	}
	return OfficeUiRibbon;
});
