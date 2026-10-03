import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CSS = `
:host { display: contents; }
.panel { position: fixed; inset: auto; margin: 0; min-width: 200px; padding: 4px; box-sizing: border-box;
	background: var(--office-background, #fff); color: var(--office-foreground, #1f2937);
	border: 1px solid var(--office-border, #d1d5db); border-radius: var(--office-radius, 4px);
	box-shadow: 0 8px 24px rgb(0 0 0 / 18%); font-family: var(--office-font, system-ui, sans-serif); }
.panel[hidden] { display: none; }
@media (forced-colors: active) { .panel { border-color: CanvasText; background: Canvas; color: CanvasText; } }
`;

const SEPARATOR_CSS = `
:host { display: block; height: 1px; margin: 4px 6px; background: var(--office-border, #d1d5db); }
@media (forced-colors: active) { :host { background: CanvasText; } }
`;

type Item = HTMLElement & { disabled: boolean };
const ITEMS = 'office-ui-menu-item:not([disabled]):not([hidden])';

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
 * Right-click menu of `office-ui-menu-item` and `office-ui-menu-separator` children (Office's
 * shape and text context menus). `openAt(x, y)` shows it at client coordinates in the popover
 * top layer, kept inside the viewport; `close()` hides it. Arrow keys, Home and End move between
 * enabled items; Escape and Tab close and return focus to the element that was focused. Choosing
 * an item closes the menu and its `office-command` bubbles to the host. `label` names the menu.
 */
export const defineContextMenu = definer('office-ui-context-menu', () => {
	class OfficeUiContextMenu extends HTMLElement {
		static observedAttributes = ['label'];
		private readonly panel: HTMLDivElement;
		private returnFocus: HTMLElement | null = null;
		private outside: ((event: Event) => void) | undefined;
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
			root.append(this.panel);
			this.panel.addEventListener('toggle', (event) => {
				if ((event as ToggleEvent).newState === 'closed') this.closed(false);
			});
			this.panel.addEventListener('keydown', (event) => this.onKey(event));
			this.addEventListener('office-command', () => this.close());
		}
		connectedCallback(): void {
			this.panel.setAttribute('aria-label', this.getAttribute('label') ?? 'Context menu');
		}
		attributeChangedCallback(): void {
			this.panel.setAttribute('aria-label', this.getAttribute('label') ?? 'Context menu');
		}
		disconnectedCallback(): void {
			this.close();
		}
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
			const width = this.panel.offsetWidth;
			const height = this.panel.offsetHeight;
			const left = view && width ? Math.min(x, view.innerWidth - width - 4) : x;
			const top = view && height ? Math.min(y, view.innerHeight - height - 4) : y;
			this.panel.style.left = `${Math.max(4, Math.round(left))}px`;
			this.panel.style.top = `${Math.max(4, Math.round(top))}px`;
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
		private onKey(event: KeyboardEvent): void {
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
			}
		}
	}
	return OfficeUiContextMenu;
});
