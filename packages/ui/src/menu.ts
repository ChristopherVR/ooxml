import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const BUTTON_CSS = `
:host { display: inline-flex; vertical-align: middle; }
.wrap { display: inline-flex; align-items: stretch; border: ${tok('--office-border-width')} solid transparent; border-radius: ${tok('--office-radius')}; }
:host([variant="stacked"]) .wrap { flex-direction: column; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center; gap: ${tok('--office-space-1')};
	min-width: ${tok('--office-target-size')}; min-height: ${tok('--office-target-size')}; padding: 0 ${tok('--office-space-1-5')};
	border: 0; border-radius: ${tok('--office-radius')}; background: transparent; color: ${tok('--office-foreground')};
	font: inherit; font-size: ${tok('--office-font-size-sm')}; cursor: pointer; white-space: nowrap; }
:host([variant="stacked"]) .main { flex-direction: column; padding: ${tok('--office-space-1')} ${tok('--office-space-1-5')} 0; }
:host([variant="stacked"]) .caret { min-height: ${tok('--office-icon-size-md')}; padding: 0 ${tok('--office-space-1-5')} ${tok('--office-space-0')}; }
:host([variant="stacked"]:not([command])) .main { padding-bottom: 0; }
.wrap:hover:not(.disabled) { border-color: ${tok('--office-border')}; }
button:hover:not(:disabled) { background: ${tok('--office-surface')}; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(-1 * ${tok('--office-focus-width')}); }
button:disabled { opacity: .5; cursor: not-allowed; }
button[aria-expanded="true"] { background: ${tok('--office-surface')}; }
.caret svg { width: ${tok('--office-icon-size-sm')}; height: ${tok('--office-icon-size-sm')}; }
svg { width: ${tok('--office-icon-size-ml')}; height: ${tok('--office-icon-size-ml')}; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
	stroke-linecap: round; stroke-linejoin: round; }
svg:not([data-painted]) { display: none; }
.panel { position: fixed; inset: auto; margin: 0; min-width: ${tok('--office-menu-min-width')}; padding: ${tok('--office-space-1')}; box-sizing: border-box;
	background: ${tok('--office-background')}; color: ${tok('--office-foreground')};
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius')};
	box-shadow: ${tok('--office-shadow')}; }
.panel[hidden] { display: none; }
@media (forced-colors: active) {
	button { color: ButtonText; } button:disabled { color: GrayText; }
	.panel { border-color: CanvasText; background: Canvas; color: CanvasText; }
}
`;

const ITEM_CSS = `
:host { display: block; }
:host([hidden]) { display: none; }
button { box-sizing: border-box; display: flex; align-items: center; gap: ${tok('--office-space-2')}; width: 100%;
	min-height: ${tok('--office-target-size')}; padding: 0 ${tok('--office-space-2-5')} 0 ${tok('--office-space-2')}; border: 0; border-radius: ${tok('--office-radius-sm')};
	background: transparent; color: inherit; font: inherit; font-size: ${tok('--office-font-size-sm')}; text-align: start; cursor: pointer; }
button:hover:not(:disabled), button:focus-visible { background: ${tok('--office-surface')}; outline: none; }
button:disabled { opacity: .5; cursor: not-allowed; }
.check { width: ${tok('--office-icon-size')}; text-align: center; }
svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; flex: none; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
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
 * and the caret opens the menu. Attributes: `label`, `icon`, `command`, `variant="stacked"`, `icon-only`,
 * `disabled`, `title`, `keyshortcuts`. The menu uses the top layer (`popover`) so ribbon
 * overflow never clips it. Keyboard: Enter, Space or ArrowDown open; arrows, Home and End move;
 * Escape and Tab close; focus returns to the trigger. Choosing an item closes the menu; its
 * `office-command` bubbles to the host.
 */
export const defineMenuButton = definer('office-ui-menu-button', () => {
	class OfficeUiMenuButton extends HTMLElement {
		static observedAttributes = [
			'label',
			'icon',
			'command',
			'disabled',
			'title',
			'keyshortcuts',
			'icon-only',
		];
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
			this.place(anchor);
			this.items()[focusIndex]?.focus();
		}
		/** Keep the open menu inside the viewport: flip left of the trigger's right edge if needed. */
		private place(anchor: DOMRect): void {
			const view = this.ownerDocument.defaultView;
			const width = this.panel.offsetWidth;
			const height = this.panel.offsetHeight;
			if (!view || !width) return;
			const maxLeft = view.innerWidth - width - 4;
			const left = anchor.left > maxLeft ? Math.max(4, anchor.right - width) : anchor.left;
			const top =
				anchor.bottom + 2 + height > view.innerHeight && anchor.top - height - 2 > 0
					? anchor.top - height - 2
					: anchor.bottom + 2;
			this.panel.style.left = `${Math.round(Math.min(left, Math.max(4, maxLeft)))}px`;
			this.panel.style.top = `${Math.round(top)}px`;
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
			const iconOnly = this.hasAttribute('icon-only');
			this.text.textContent = label;
			this.text.hidden = iconOnly;
			if (iconOnly) this.main.setAttribute('aria-label', label);
			else this.main.removeAttribute('aria-label');
			const title = this.getAttribute('title') ?? label;
			this.main.title = title;
			this.caret.title = split ? `${label} options` : title;
			// Without a split command the caret is decoration on one control, not a second button.
			if (split) {
				this.caret.setAttribute('aria-label', `${label} options`);
				this.caret.removeAttribute('aria-hidden');
			} else {
				this.caret.removeAttribute('aria-label');
				this.caret.setAttribute('aria-hidden', 'true');
			}
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
