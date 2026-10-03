import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const BUTTON_CSS = `
:host { display: inline-flex; vertical-align: middle; }
.wrap { display: inline-flex; align-items: stretch; border: 1px solid transparent; border-radius: var(--office-radius, 4px); }
:host([variant="stacked"]) .wrap { flex-direction: column; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center; gap: 4px;
	min-width: var(--office-target-size, 28px); min-height: var(--office-target-size, 28px); padding: 0 6px;
	border: 0; border-radius: var(--office-radius, 4px); background: transparent; color: var(--office-foreground, #1f2937);
	font: inherit; font-size: 12px; cursor: pointer; white-space: nowrap; }
:host([variant="stacked"]) .main { flex-direction: column; padding: 4px 6px 0; }
:host([variant="stacked"]) .caret { min-height: 16px; padding: 0 6px 2px; }
:host([variant="stacked"]:not([command])) .main { padding-bottom: 0; }
.wrap:hover:not(.disabled) { border-color: var(--office-border, #d1d5db); }
button:hover:not(:disabled) { background: var(--office-surface, #f3f4f6); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: -2px; }
button:disabled { opacity: .5; cursor: not-allowed; }
button[aria-expanded="true"] { background: var(--office-surface, #f3f4f6); }
.caret svg { width: 12px; height: 12px; }
svg { width: 18px; height: 18px; fill: none; stroke: currentColor; stroke-width: 1.5;
	stroke-linecap: round; stroke-linejoin: round; }
svg:not([data-painted]) { display: none; }
.panel { position: fixed; inset: auto; margin: 0; min-width: 180px; padding: 4px; box-sizing: border-box;
	background: var(--office-background, #fff); color: var(--office-foreground, #1f2937);
	border: 1px solid var(--office-border, #d1d5db); border-radius: var(--office-radius, 4px);
	box-shadow: 0 8px 24px rgb(0 0 0 / 18%); }
.panel[hidden] { display: none; }
@media (forced-colors: active) {
	button { color: ButtonText; } button:disabled { color: GrayText; }
	.panel { border-color: CanvasText; background: Canvas; color: CanvasText; }
}
`;

const ITEM_CSS = `
:host { display: block; }
:host([hidden]) { display: none; }
button { box-sizing: border-box; display: flex; align-items: center; gap: 8px; width: 100%;
	min-height: var(--office-target-size, 28px); padding: 0 10px 0 8px; border: 0; border-radius: 3px;
	background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: start; cursor: pointer; }
button:hover:not(:disabled), button:focus-visible { background: var(--office-surface, #f3f4f6); outline: none; }
button:disabled { opacity: .5; cursor: not-allowed; }
.check { width: 14px; text-align: center; }
svg { width: 16px; height: 16px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.5;
	stroke-linecap: round; stroke-linejoin: round; }
svg:not([data-painted]) { visibility: hidden; }
@media (forced-colors: active) { button:focus-visible { background: Highlight; color: HighlightText; } }
`;

type Item = HTMLElement & { disabled: boolean };
const ENABLED_ITEMS = 'office-ui-menu-item:not([disabled]):not([hidden])';

/**
 * One menu entry: `command`, `label`, `icon`, `disabled`, `checked` (renders a
 * `menuitemcheckbox`), `title`. Activation emits `office-command` `{ command }`.
 */
export const defineMenuItem = definer('office-ui-menu-item', () => {
	class OfficeUiMenuItem extends HTMLElement {
		static observedAttributes = ['label', 'icon', 'command', 'disabled', 'checked', 'title'];
		private readonly button: HTMLButtonElement;
		private readonly svg: SVGSVGElement;
		private readonly text: HTMLSpanElement;
		private readonly mark: HTMLSpanElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(ITEM_CSS));
			this.button = doc.createElement('button');
			this.button.type = 'button';
			this.button.tabIndex = -1;
			this.mark = doc.createElement('span');
			this.mark.className = 'check';
			this.mark.setAttribute('aria-hidden', 'true');
			this.svg = createIconSvg(doc);
			this.text = doc.createElement('span');
			this.button.append(this.mark, this.svg, this.text);
			root.append(this.button);
			this.button.addEventListener('click', () => {
				const command = this.getAttribute('command');
				if (!this.button.disabled && command) emit(this, 'office-command', { command });
			});
		}
		connectedCallback(): void {
			this.sync();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(value: boolean) {
			this.toggleAttribute('disabled', Boolean(value));
		}
		/** Focus the menuitem itself, so arrow navigation never depends on delegatesFocus. */
		override focus(options?: FocusOptions): void {
			this.button.focus(options);
		}
		private sync(): void {
			const checked = this.getAttribute('checked');
			this.button.setAttribute('role', checked === null ? 'menuitem' : 'menuitemcheckbox');
			if (checked === null) this.button.removeAttribute('aria-checked');
			else this.button.setAttribute('aria-checked', String(checked !== 'false'));
			this.mark.textContent = checked !== null && checked !== 'false' ? '✓' : '';
			this.text.textContent = this.getAttribute('label') ?? '';
			this.button.title = this.getAttribute('title') ?? '';
			this.button.disabled = this.disabled;
			if (paintIcon(this.svg, this.getAttribute('icon'))) this.svg.setAttribute('data-painted', '');
			else this.svg.removeAttribute('data-painted');
		}
	}
	return OfficeUiMenuItem;
});

/**
 * Dropdown command with a menu of `office-ui-menu-item` children (Office's Find, Layers,
 * Align). With a `command` attribute it is a split button: the main part emits that command
 * and the caret opens the menu. Attributes: `label`, `icon`, `command`, `variant="stacked"`,
 * `disabled`, `title`, `keyshortcuts`. The menu uses the top layer (`popover`) so ribbon
 * overflow never clips it. Keyboard: Enter, Space or ArrowDown open; arrows, Home and End move;
 * Escape and Tab close; focus returns to the trigger. Choosing an item closes the menu; its
 * `office-command` bubbles to the host.
 */
export const defineMenuButton = definer('office-ui-menu-button', () => {
	class OfficeUiMenuButton extends HTMLElement {
		static observedAttributes = ['label', 'icon', 'command', 'disabled', 'title', 'keyshortcuts'];
		private readonly wrap: HTMLDivElement;
		private readonly main: HTMLButtonElement;
		private readonly caret: HTMLButtonElement;
		private readonly svg: SVGSVGElement;
		private readonly text: HTMLSpanElement;
		private readonly panel: HTMLDivElement;
		private outside: ((event: Event) => void) | undefined;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(BUTTON_CSS));
			this.wrap = doc.createElement('div');
			this.wrap.className = 'wrap';
			this.main = doc.createElement('button');
			this.main.type = 'button';
			this.main.className = 'main';
			this.main.setAttribute('part', 'button');
			this.svg = createIconSvg(doc);
			this.text = doc.createElement('span');
			this.main.append(this.svg, this.text);
			this.caret = doc.createElement('button');
			this.caret.type = 'button';
			this.caret.className = 'caret';
			this.caret.setAttribute('aria-haspopup', 'menu');
			this.caret.setAttribute('aria-expanded', 'false');
			const chevron = createIconSvg(doc);
			paintIcon(chevron, 'chevronDown');
			chevron.setAttribute('data-painted', '');
			this.caret.append(chevron);
			this.panel = doc.createElement('div');
			this.panel.className = 'panel';
			this.panel.setAttribute('role', 'menu');
			if (typeof (this.panel as { showPopover?: unknown }).showPopover === 'function')
				this.panel.setAttribute('popover', 'auto');
			else this.panel.hidden = true;
			this.panel.append(doc.createElement('slot'));
			this.wrap.append(this.main, this.caret);
			root.append(this.wrap, this.panel);
			this.main.addEventListener('click', () => {
				const command = this.getAttribute('command');
				if (this.disabled) return;
				if (command) emit(this, 'office-command', { command });
				else this.toggle();
			});
			this.caret.addEventListener('click', () => {
				if (!this.disabled) this.toggle();
			});
			for (const trigger of [this.main, this.caret])
				trigger.addEventListener('keydown', (event) => {
					if (event.key === 'ArrowDown' && !this.disabled) {
						event.preventDefault();
						this.show(0);
					}
				});
			this.panel.addEventListener('toggle', (event) => {
				if ((event as ToggleEvent).newState === 'closed') this.closed();
			});
			this.panel.addEventListener('keydown', (event) => this.onMenuKey(event));
			// Item choices bubble on to the host; the menu only closes.
			this.addEventListener('office-command', (event) => {
				if (event.target !== this) this.hide(true);
			});
		}
		connectedCallback(): void {
			this.sync();
		}
		disconnectedCallback(): void {
			this.hide(false);
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(value: boolean) {
			this.toggleAttribute('disabled', Boolean(value));
		}
		get open(): boolean {
			return this.caret.getAttribute('aria-expanded') === 'true';
		}
		private items(): Item[] {
			return [...this.querySelectorAll<Item>(ENABLED_ITEMS)];
		}
		private toggle(): void {
			if (this.open) this.hide(true);
			else this.show(0);
		}
		private show(focusIndex: number): void {
			const anchor = this.wrap.getBoundingClientRect();
			this.panel.style.left = `${Math.round(anchor.left)}px`;
			this.panel.style.top = `${Math.round(anchor.bottom + 2)}px`;
			if (this.panel.hasAttribute('popover')) this.panel.showPopover?.();
			else {
				this.panel.hidden = false;
				this.outside = (event) => {
					if (!event.composedPath().includes(this)) this.hide(false);
				};
				this.ownerDocument.addEventListener('pointerdown', this.outside, true);
			}
			this.setExpanded(true);
			this.items()[focusIndex]?.focus();
		}
		private hide(restoreFocus: boolean): void {
			if (!this.open) return;
			if (this.panel.hasAttribute('popover')) {
				try {
					this.panel.hidePopover?.();
				} catch {
					/* Already closed by light dismiss. */
				}
			}
			this.closed();
			if (restoreFocus) (this.getAttribute('command') ? this.caret : this.main).focus();
		}
		private closed(): void {
			if (!this.panel.hasAttribute('popover')) this.panel.hidden = true;
			if (this.outside) this.ownerDocument.removeEventListener('pointerdown', this.outside, true);
			this.outside = undefined;
			this.setExpanded(false);
		}
		private setExpanded(open: boolean): void {
			this.caret.setAttribute('aria-expanded', String(open));
			if (this.hasAttribute('command')) this.main.removeAttribute('aria-expanded');
			else this.main.setAttribute('aria-expanded', String(open));
		}
		private onMenuKey(event: KeyboardEvent): void {
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
			else if (event.key === 'Escape') {
				event.preventDefault();
				this.hide(true);
			} else if (event.key === 'Tab') this.hide(false);
		}
		private sync(): void {
			const label = this.getAttribute('label') ?? '';
			const split = this.hasAttribute('command');
			this.text.textContent = label;
			const title = this.getAttribute('title') ?? label;
			this.main.title = title;
			this.caret.title = split ? `${label} options` : title;
			this.caret.setAttribute('aria-label', split ? `${label} options` : label);
			if (split) {
				this.main.removeAttribute('aria-haspopup');
				this.main.removeAttribute('aria-expanded');
				this.caret.tabIndex = 0;
			} else {
				// Without a command the whole control opens the menu; the caret is decorative.
				this.main.setAttribute('aria-haspopup', 'menu');
				this.main.setAttribute('aria-expanded', String(this.open));
				this.caret.tabIndex = -1;
			}
			const shortcuts = this.getAttribute('keyshortcuts');
			if (shortcuts) this.main.setAttribute('aria-keyshortcuts', shortcuts);
			else this.main.removeAttribute('aria-keyshortcuts');
			this.main.disabled = this.caret.disabled = this.disabled;
			this.wrap.classList.toggle('disabled', this.disabled);
			if (this.disabled) this.hide(false);
			if (paintIcon(this.svg, this.getAttribute('icon'))) this.svg.setAttribute('data-painted', '');
			else this.svg.removeAttribute('data-painted');
		}
	}
	return OfficeUiMenuButton;
});
