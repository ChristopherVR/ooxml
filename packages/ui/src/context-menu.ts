import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import {
	clampFlyoutPosition,
	EMPTY_MENU_STATE,
	nextEnabledIndex,
	typeAheadIndex,
	type OfficeMenuCloseReason,
	type OfficeMenuState,
} from './menu-model.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CSS = `
:host { display: contents; }
:host([data-controlled]) { display: block; position: fixed; z-index: ${tok('--office-z-popover')};
	color: ${tok('--office-popover-foreground')}; font: ${tok('--office-font-size-sm')}/${tok('--office-line-height')} ${tok('--office-font')}; }
:host([data-controlled][hidden]) { display: none; }
.panel { position: fixed; inset: auto; margin: 0; min-width: ${tok('--office-context-menu-min-width')}; padding: ${tok('--office-space-1')}; box-sizing: border-box;
	background: ${tok('--office-background')}; color: ${tok('--office-foreground')};
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius')};
	box-shadow: ${tok('--office-shadow')}; font-family: ${tok('--office-font')}; }
.panel[hidden] { display: none; }
/* Controlled mode: the host is the fixed box; rows are drawn here from \`state\`. */
.menu { box-sizing: border-box; display: flex; flex-direction: column; min-width: ${tok('--office-menu-min-width')};
	max-height: calc(100vh - 2 * ${tok('--office-space-2')}); overflow-y: auto; padding: ${tok('--office-space-1-5')} 0;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-md')};
	background: ${tok('--office-popover')}; color: inherit; box-shadow: ${tok('--office-shadow-lg')}; user-select: none; }
.menu:focus { outline: none; }
.item { box-sizing: border-box; display: flex; align-items: center; gap: ${tok('--office-space-1-5')}; width: 100%;
	min-height: ${tok('--office-field-height')}; padding: ${tok('--office-space-1-5')} ${tok('--office-space-3')}; border: 0;
	background: transparent; color: inherit; font: inherit; text-align: start; white-space: nowrap; cursor: pointer; }
.item:hover:not(:disabled), .item:focus:not(:disabled) { background: ${tok('--office-selected')}; outline: none; }
.item:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(-1 * ${tok('--office-focus-width')}); }
.item:disabled { opacity: .45; cursor: default; }
.item.danger { color: ${tok('--office-danger')}; }
.item svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; flex: none; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
.check { flex: none; width: ${tok('--office-icon-size')}; text-align: center; }
.separator { height: ${tok('--office-border-width')}; margin: ${tok('--office-space-1')} 0; background: ${tok('--office-border')}; }
.heading { padding: ${tok('--office-space-1-5')} ${tok('--office-space-3')} ${tok('--office-space-0')}; color: ${tok('--office-muted-foreground')};
	font-size: ${tok('--office-font-size-2xs')}; font-weight: ${tok('--office-font-weight-bold')}; letter-spacing: .06em; text-transform: uppercase; }
@media (pointer: coarse) { .item { min-height: ${tok('--office-target-size-touch')}; } }
@media (forced-colors: active) {
	.panel, .menu { border-color: CanvasText; background: Canvas; color: CanvasText; box-shadow: none; }
	.item, .item.danger { color: CanvasText; }
	.item:disabled { color: GrayText; opacity: 1; }
	.item:hover:not(:disabled), .item:focus:not(:disabled) { background: Highlight; color: HighlightText; }
	.separator { background: CanvasText; }
	.heading { color: CanvasText; }
}
`;

const SEPARATOR_CSS = `
:host { display: block; height: ${tok('--office-border-width')}; margin: ${tok('--office-space-1')} ${tok('--office-space-1-5')}; background: ${tok('--office-border')}; }
@media (forced-colors: active) { :host { background: CanvasText; } }
`;

type Item = HTMLElement & { disabled: boolean };
const ITEMS = 'office-ui-menu-item:not([disabled]):not([hidden])';
const TYPE_AHEAD_MS = 700;
type MenuConfig = { requestEvent: string; closeEvent: string; markerPattern: RegExp };

/** Open controlled menus per document, innermost last: Escape closes only the top one. */
const openMenus = new WeakMap<Document, HTMLElement[]>();

/** A `role="separator"` rule between groups of menu items. */
export const defineMenuSeparator = definer('office-ui-menu-separator', () => {
	class OfficeUiMenuSeparator extends HTMLElement {
		constructor() {
			super();
			attachStyles(this.attachShadow({ mode: 'open' }), controlCss(SEPARATOR_CSS));
		}
		connectedCallback(): void {
			this.setAttribute('role', 'separator');
		}
	}
	return OfficeUiMenuSeparator;
});

/**
 * Office's right-click menu, in two modes sharing keyboard, type-ahead and dismissal:
 *
 * - **Children** (self-managed): `office-ui-menu-item` and `office-ui-menu-separator` children.
 *   `openAt(x, y)` shows it in the popover top layer, kept inside the viewport; `close()` hides it.
 *   Choosing an item closes the menu and its `office-command` bubbles to the host.
 * - **Controlled** (moved from pptx-viewer `pptx-ui-context-menu`): set `state`
 *   (`{ x, y, label, items: [{ id, label, separatorBefore?, heading?, danger?, disabled?,
 *   checked?, icon? }], markers?, zIndex?, autoFocus? }`). The host is the fixed box at the
 *   pointer, clamped into the window. It emits `office-menu-request` `{ id }` when a row is chosen
 *   and `office-menu-close` `{ reason: 'escape' | 'outside' | 'tab' }` when dismissed, and never
 *   closes itself: the host removes it or clears its items. While open it takes focus on the first
 *   enabled row and returns focus to the opener when it goes away, unless something else took it.
 *
 * Arrow keys, Home, End and type-ahead (repeating a letter cycles) move between enabled rows.
 * Static `requestEvent`, `closeEvent` and `markerPattern` let a product keep its published names.
 */
export const defineContextMenu = definer('office-ui-context-menu', () => {
	class OfficeUiContextMenu extends HTMLElement {
		static requestEvent = 'office-menu-request';
		static closeEvent = 'office-menu-close';
		static markerPattern = /^data-[a-z0-9-]+$/;
		static observedAttributes = ['label'];
		private readonly panel: HTMLDivElement;
		private readonly menu: HTMLDivElement;
		private returnFocus: HTMLElement | null = null;
		private outside: ((event: Event) => void) | undefined;
		private model: OfficeMenuState = EMPTY_MENU_STATE;
		private controlled = false;
		private rows: HTMLButtonElement[] = [];
		private active = -1;
		private focused = false;
		private query = '';
		private queryTimer: ReturnType<typeof setTimeout> | undefined;
		private markers: string[] = [];
		private listening = false;
		private signature = '';
		private opener: HTMLElement | null = null;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.panel = doc.createElement('div');
			this.panel.className = 'panel';
			this.panel.setAttribute('role', 'menu');
			if (typeof (this.panel as { showPopover?: unknown }).showPopover === 'function')
				this.panel.setAttribute('popover', 'auto');
			else this.panel.hidden = true;
			this.panel.append(doc.createElement('slot'));
			this.menu = doc.createElement('div');
			this.menu.className = 'menu';
			this.menu.setAttribute('role', 'menu');
			this.menu.setAttribute('aria-orientation', 'vertical');
			this.menu.tabIndex = -1;
			this.menu.hidden = true;
			root.append(this.panel, this.menu);
			this.panel.addEventListener('toggle', (event) => {
				if ((event as ToggleEvent).newState === 'closed') this.closed(false);
			});
			this.panel.addEventListener('keydown', (event) => this.onChildKey(event));
			this.menu.addEventListener('keydown', (event) => this.onRowKey(event));
			this.menu.addEventListener('pointerover', (event) => {
				const row = (event.target as Element | null)?.closest('button');
				const index = row ? this.rows.indexOf(row) : -1;
				if (index >= 0 && !row?.disabled && index !== this.active) this.focusRow(index);
			});
			this.menu.addEventListener('contextmenu', (event) => event.preventDefault());
			this.addEventListener('office-command', () => this.close());
		}
		connectedCallback(): void {
			this.panel.setAttribute('aria-label', this.getAttribute('label') ?? 'Context menu');
			if (this.controlled) {
				this.opener = this.deepActive();
				this.render();
			}
		}
		attributeChangedCallback(): void {
			this.panel.setAttribute('aria-label', this.getAttribute('label') ?? 'Context menu');
		}
		disconnectedCallback(): void {
			this.close();
			if (!this.controlled) return;
			this.stopListening();
			clearTimeout(this.queryTimer);
			this.focused = false;
			const opener = this.opener;
			this.opener = null;
			// Only restore when nothing else took focus (a command may open a dialog or editor).
			const active = this.ownerDocument.activeElement;
			if (opener?.isConnected && (!active || active === this.ownerDocument.body))
				opener.focus({ preventScroll: true });
		}

		// ---- Children mode ------------------------------------------------------------------
		get open(): boolean {
			return this.hasAttribute('data-open');
		}
		private items(): Item[] {
			return [...this.querySelectorAll<Item>(ITEMS)];
		}
		/** Show at client coordinates; focus the first enabled item. */
		openAt(x: number, y: number): void {
			const doc = this.ownerDocument;
			const active = doc.activeElement;
			this.returnFocus = active instanceof HTMLElement ? active : null;
			if (this.panel.hasAttribute('popover')) {
				if (!this.open) this.panel.showPopover?.();
			} else {
				this.panel.hidden = false;
				this.outside ??= (event) => {
					if (!event.composedPath().includes(this)) this.close();
				};
				doc.addEventListener('pointerdown', this.outside, true);
			}
			this.setAttribute('data-open', '');
			const view = doc.defaultView;
			const { left, top } = clampFlyoutPosition({
				x,
				y,
				width: this.panel.offsetWidth,
				height: this.panel.offsetHeight,
				viewportWidth: view?.innerWidth ?? x,
				viewportHeight: view?.innerHeight ?? y,
				margin: 4,
			});
			this.panel.style.left = `${Math.round(left)}px`;
			this.panel.style.top = `${Math.round(top)}px`;
			this.items()[0]?.focus();
		}
		close(): void {
			if (!this.open) return;
			if (this.panel.hasAttribute('popover')) {
				try {
					this.panel.hidePopover?.();
				} catch {
					/* Already closed by light dismiss. */
				}
			}
			this.closed(true);
		}
		private closed(restoreFocus: boolean): void {
			if (!this.open) return;
			this.removeAttribute('data-open');
			if (!this.panel.hasAttribute('popover')) this.panel.hidden = true;
			if (this.outside) this.ownerDocument.removeEventListener('pointerdown', this.outside, true);
			if (restoreFocus) this.returnFocus?.focus?.();
			this.returnFocus = null;
		}
		private onChildKey(event: KeyboardEvent): void {
			const items = this.items();
			const from = items.findIndex(
				(item) => item === event.target || item.contains(event.target as Node),
			);
			const move = (index: number) => {
				event.preventDefault();
				items[(index + items.length) % items.length]?.focus();
			};
			if (event.key === 'ArrowDown') move(from + 1);
			else if (event.key === 'ArrowUp') move(from - 1);
			else if (event.key === 'Home') move(0);
			else if (event.key === 'End') move(items.length - 1);
			else if (event.key === 'Escape' || event.key === 'Tab') {
				if (event.key === 'Escape') event.preventDefault();
				this.close();
			} else if (
				event.key.length === 1 &&
				event.key !== ' ' &&
				!event.ctrlKey &&
				!event.metaKey &&
				!event.altKey
			) {
				const labels = items.map((item) => ({
					label: item.getAttribute('label') ?? item.textContent ?? '',
				}));
				const target = this.typeAhead(labels, from, event.key);
				if (target >= 0) move(target);
			}
		}

		// ---- Controlled mode ----------------------------------------------------------------
		get state(): OfficeMenuState {
			return this.model;
		}
		set state(value: OfficeMenuState) {
			const signature = JSON.stringify(value);
			this.model = value;
			if (!this.controlled) {
				this.controlled = true;
				this.toggleAttribute('data-controlled', true);
				// The children-mode panel would be a second menu for assistive technology.
				this.panel.remove();
				if (this.isConnected) this.opener = this.deepActive();
			}
			// Hosts re-send equal state on unrelated renders; rebuilding rows would drop hover.
			if (signature !== this.signature) {
				this.signature = signature;
				this.render();
			}
		}
		private config(): MenuConfig {
			return this.constructor as unknown as MenuConfig;
		}
		private deepActive(): HTMLElement | null {
			let node: Element | null = this.ownerDocument.activeElement;
			while (node?.shadowRoot?.activeElement) node = node.shadowRoot.activeElement;
			return node instanceof HTMLElement && node !== this.ownerDocument.body ? node : null;
		}
		private render(): void {
			const state = this.model;
			const items = state.items;
			this.hidden = items.length === 0;
			this.menu.hidden = items.length === 0;
			this.applyMarkers(state.markers ?? []);
			if (state.zIndex !== undefined) this.style.zIndex = String(state.zIndex);
			else this.style.removeProperty('z-index');
			const same = items[this.active];
			this.active = same && !same.disabled ? this.active : nextEnabledIndex(items, -1, 1);
			this.renderRows();
			if (items.length === 0) {
				this.stopListening();
				this.focused = false;
			}
			if (!this.isConnected || items.length === 0) return;
			this.place();
			this.listen();
			if (!this.focused && state.autoFocus !== false) {
				this.focused = true;
				this.focusRow(nextEnabledIndex(items, -1, 1));
			} else if (this.focused && this.active >= 0)
				this.rows[this.active]?.focus({ preventScroll: true });
		}
		private renderRows(): void {
			const doc = this.ownerDocument;
			const state = this.model;
			this.menu.setAttribute('aria-label', state.label);
			this.rows = [];
			let container: HTMLElement = this.menu;
			const nodes: HTMLElement[] = [];
			const place = (node: HTMLElement) =>
				container === this.menu ? nodes.push(node) : container.append(node);
			state.items.forEach((item, index) => {
				if (item.separatorBefore || item.heading) container = this.menu;
				if (item.separatorBefore) {
					const rule = doc.createElement('div');
					rule.className = 'separator';
					rule.setAttribute('role', 'separator');
					place(rule);
				}
				if (item.heading) {
					const group = doc.createElement('div');
					group.setAttribute('role', 'group');
					group.setAttribute('aria-label', item.heading);
					const title = doc.createElement('div');
					title.className = 'heading';
					title.setAttribute('aria-hidden', 'true');
					title.textContent = item.heading;
					group.append(title);
					place(group);
					container = group;
				}
				const row = doc.createElement('button');
				row.type = 'button';
				row.className = item.danger ? 'item danger' : 'item';
				row.dataset.itemId = item.id;
				row.setAttribute('role', item.checked === undefined ? 'menuitem' : 'menuitemcheckbox');
				if (item.checked !== undefined) {
					row.setAttribute('aria-checked', String(item.checked));
					const check = doc.createElement('span');
					check.className = 'check';
					check.setAttribute('aria-hidden', 'true');
					check.textContent = item.checked ? '✓' : '';
					row.append(check);
				}
				if (item.icon) {
					const glyph = createIconSvg(doc);
					paintIcon(glyph, item.icon);
					row.append(glyph);
				}
				row.append(doc.createTextNode(item.label));
				row.tabIndex = index === this.active ? 0 : -1;
				if (item.disabled) {
					row.disabled = true;
					row.setAttribute('aria-disabled', 'true');
				}
				row.addEventListener('click', () =>
					this.dispatchEvent(
						new CustomEvent(this.config().requestEvent, {
							detail: { id: item.id },
							bubbles: true,
							composed: true,
						}),
					),
				);
				this.rows.push(row);
				place(row);
			});
			this.menu.replaceChildren(...nodes);
		}
		private applyMarkers(next: readonly string[]): void {
			for (const name of this.markers) this.removeAttribute(name);
			const pattern = this.config().markerPattern;
			this.markers = next.filter((name) => pattern.test(name));
			for (const name of this.markers) this.setAttribute(name, 'true');
		}
		private place(): void {
			const view = this.ownerDocument.defaultView;
			this.style.left = `${this.model.x}px`;
			this.style.top = `${this.model.y}px`;
			const box = this.getBoundingClientRect();
			const { left, top } = clampFlyoutPosition({
				x: this.model.x,
				y: this.model.y,
				width: box.width,
				height: box.height,
				viewportWidth: view?.innerWidth ?? box.right,
				viewportHeight: view?.innerHeight ?? box.bottom,
			});
			this.style.left = `${left}px`;
			this.style.top = `${top}px`;
		}
		private focusRow(index: number): void {
			if (index < 0) return;
			this.active = index;
			this.rows.forEach((row, at) => (row.tabIndex = at === index ? 0 : -1));
			this.rows[index]?.focus({ preventScroll: true });
		}
		private typeAhead(
			items: readonly { label: string; disabled?: boolean }[],
			from: number,
			key: string,
		): number {
			this.query += key;
			clearTimeout(this.queryTimer);
			this.queryTimer = setTimeout(() => (this.query = ''), TYPE_AHEAD_MS);
			return typeAheadIndex(items, from, this.query);
		}
		private onRowKey(event: KeyboardEvent): void {
			const items = this.model.items;
			let target = -2;
			if (event.key === 'ArrowDown') target = nextEnabledIndex(items, this.active, 1);
			else if (event.key === 'ArrowUp') target = nextEnabledIndex(items, this.active, -1);
			else if (event.key === 'Home') target = nextEnabledIndex(items, -1, 1);
			else if (event.key === 'End') target = nextEnabledIndex(items, items.length, -1);
			else if (
				event.key.length === 1 &&
				event.key !== ' ' &&
				!event.ctrlKey &&
				!event.metaKey &&
				!event.altKey
			)
				target = this.typeAhead(items, this.active, event.key);
			if (target === -2) return;
			event.preventDefault();
			this.focusRow(target);
		}
		private dismiss(reason: OfficeMenuCloseReason): void {
			this.dispatchEvent(
				new CustomEvent(this.config().closeEvent, {
					detail: { reason },
					bubbles: true,
					composed: true,
				}),
			);
		}
		private readonly onDocumentKey = (event: KeyboardEvent): void => {
			if (openMenus.get(this.ownerDocument)?.at(-1) !== this) return;
			if (event.key === 'Escape') {
				// The menu consumes Escape so a full-screen view behind it does not also exit.
				event.preventDefault();
				event.stopPropagation();
				this.dismiss('escape');
			} else if (event.key === 'Tab') this.dismiss('tab');
		};
		private readonly onDocumentPointer = (event: Event): void => {
			if (!event.composedPath().includes(this)) this.dismiss('outside');
		};
		private listen(): void {
			if (this.listening) return;
			this.listening = true;
			const doc = this.ownerDocument;
			openMenus.set(doc, [...(openMenus.get(doc) ?? []), this]);
			doc.addEventListener('keydown', this.onDocumentKey, true);
			doc.addEventListener('pointerdown', this.onDocumentPointer, true);
		}
		private stopListening(): void {
			if (!this.listening) return;
			this.listening = false;
			const doc = this.ownerDocument;
			openMenus.set(
				doc,
				(openMenus.get(doc) ?? []).filter((menu) => menu !== this),
			);
			doc.removeEventListener('keydown', this.onDocumentKey, true);
			doc.removeEventListener('pointerdown', this.onDocumentPointer, true);
		}
	}
	return OfficeUiContextMenu;
});
