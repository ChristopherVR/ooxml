import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeCommandEvent = CustomEvent<{ command: string }>;

const CSS = `
:host { display: inline-flex; vertical-align: middle; }
button { box-sizing: border-box; position: relative; display: inline-flex; align-items: center; justify-content: center;
	gap: ${tok('--office-space-1-5')}; min-width: ${tok('--office-target-size')}; min-height: ${tok('--office-target-size')};
	padding: 0 ${tok('--office-space-2')}; border: ${tok('--office-border-width')} solid transparent; border-radius: ${tok('--office-radius')};
	background: transparent; color: ${tok('--office-foreground')}; font: inherit; font-size: ${tok('--office-font-size-sm')};
	cursor: pointer; white-space: nowrap; }
:host([variant="stacked"]) button { flex-direction: column; padding: ${tok('--office-space-1')} ${tok('--office-space-2')}; }
button:hover:not(:disabled) { background: ${tok('--office-surface')}; }
button:active:not(:disabled) { background: ${tok('--office-selected')}; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
button[aria-pressed="true"], button[aria-expanded="true"] { border-color: ${tok('--office-accent')};
	background: ${tok('--office-surface')}; }
:host([active]) button { background: ${tok('--office-command-active')}; color: ${tok('--office-accent')}; }
button:disabled { opacity: .5; cursor: not-allowed; }
svg { width: ${tok('--office-icon-size-ml')}; height: ${tok('--office-icon-size-ml')}; flex: none; color: ${tok('--office-command-icon-color')};
	fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
svg.glyph:not([data-painted]) { display: none; }
.label[hidden], .caret[hidden], .badge[hidden] { display: none; }
/* Office trails a menu chevron after the last line of the label. */
.caret { display: inline-block; width: ${tok('--office-icon-size-xs')}; height: ${tok('--office-icon-size-xs')};
	margin-inline-start: ${tok('--office-space-0')}; vertical-align: middle; color: ${tok('--office-muted-foreground')}; }
.badge { position: absolute; top: 0; inset-inline-end: 0; padding: 0 ${tok('--office-space-0')}; border-radius: ${tok('--office-radius-full')};
	background: ${tok('--office-accent')}; color: ${tok('--office-accent-foreground')}; font-size: ${tok('--office-font-size-3xs')}; }
/* Large ribbon command: a big glyph above a two-line label. */
:host([size="large"]) button { flex-direction: column; justify-content: flex-start; gap: ${tok('--office-space-0')};
	min-width: ${tok('--office-command-large-min-width')}; max-width: ${tok('--office-command-large-max-width')};
	height: ${tok('--office-command-large-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-1')};
	border: 0; white-space: normal; text-align: center; line-height: ${tok('--office-line-height-tight')}; }
:host([size="large"]) svg.glyph { width: ${tok('--office-command-large-icon')}; height: ${tok('--office-command-large-icon')}; }
/* Small ribbon command: icon and label in one row. */
:host([size="small"]) button { justify-content: flex-start; min-width: 0; height: ${tok('--office-control-height')};
	min-height: ${tok('--office-control-height')}; padding: 0 ${tok('--office-space-1-5')}; border: 0; text-align: start; }
:host([size="small"]) svg.glyph { width: ${tok('--office-command-small-icon')}; height: ${tok('--office-command-small-icon')}; }
/* Tall, label-less tool tiles (pens, erasers). */
:host([icon-only][tall]) button { min-width: ${tok('--office-command-tall-min-width')}; height: ${tok('--office-command-tall-height')}; padding: ${tok('--office-space-1')}; }
:host([icon-only][tall]) svg.glyph { width: ${tok('--office-command-tall-icon')}; height: ${tok('--office-command-tall-icon')}; }
:host([icon-only][tall][active]) button { box-shadow: inset 0 0 0 ${tok('--office-border-width')} ${tok('--office-accent')}; }
@media (pointer: coarse), (max-width: 767px) {
	:host([size="small"]) button { min-height: ${tok('--office-target-size-touch')}; }
	:host([icon-only]) button { min-width: ${tok('--office-target-size-touch')}; }
}
@media (forced-colors: active) {
	button { border-color: ButtonText; color: ButtonText; background: ButtonFace; }
	svg { color: ButtonText; }
	button:disabled { color: GrayText; border-color: GrayText; }
	button[aria-pressed="true"], button[aria-expanded="true"], :host([active]) button { background: Highlight; color: HighlightText; }
}
`;

type Configured = { requestEvent: string; idAttribute: string; detailKey: string };

/**
 * Command button. Attributes: `label`, `icon`, `command`, `disabled`, `pressed` (aria-pressed),
 * `expanded` (aria-expanded), `icon-only`, `variant="stacked"`, `size="large|small"` (ribbon
 * large and small commands), `tall` (with `icon-only`: a tool tile), `active`, `badge`, `caret`
 * (a trailing menu chevron), `title`, `keyshortcuts` (forwarded as aria-keyshortcuts).
 * Activation (pointer, Enter, Space) emits one bubbling, composed `office-command` `{ command }`;
 * Enter and Space never reach the host's own key handlers. Nothing happens when disabled.
 * Static `requestEvent`, `idAttribute` and `detailKey`, and the `iconName()` and `showsCaret()`
 * methods, let a product subclass keep its published contract and glyph set.
 */
export const defineButton = definer('office-ui-button', () => {
	class OfficeUiButton extends HTMLElement {
		static requestEvent = 'office-command';
		static idAttribute = 'command';
		static detailKey = 'command';
		static get observedAttributes(): string[] {
			return [
				'label',
				'icon',
				'command',
				'disabled',
				'pressed',
				'expanded',
				'icon-only',
				'title',
				'keyshortcuts',
				'badge',
				'caret',
				(this as unknown as Configured).idAttribute,
			];
		}
		private readonly button: HTMLButtonElement;
		private readonly svg: SVGSVGElement;
		private readonly text: HTMLSpanElement;
		private readonly caret: SVGSVGElement;
		private readonly badge: HTMLSpanElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			this.button = doc.createElement('button');
			this.button.type = 'button';
			this.button.setAttribute('part', 'button');
			this.svg = createIconSvg(doc);
			this.svg.classList.add('glyph');
			this.text = doc.createElement('span');
			this.text.className = 'label';
			this.caret = createIconSvg(doc);
			this.caret.classList.add('caret');
			paintIcon(this.caret, 'chevronDown');
			this.badge = doc.createElement('span');
			this.badge.className = 'badge';
			this.badge.setAttribute('aria-hidden', 'true');
			this.button.append(this.svg, this.text, this.badge);
			root.append(this.button);
			this.button.addEventListener('keydown', (event) => {
				// Keep native activation out of the host's own shortcut handlers.
				if (
					(event.key === ' ' || event.key === 'Enter') &&
					!event.ctrlKey &&
					!event.metaKey &&
					!event.altKey
				)
					event.stopPropagation();
			});
			this.button.addEventListener('click', () => {
				const { requestEvent, idAttribute, detailKey } = this.constructor as unknown as Configured;
				const id = this.getAttribute(idAttribute);
				if (!this.button.disabled && id)
					this.dispatchEvent(
						new CustomEvent(requestEvent, {
							detail: { [detailKey]: id },
							bubbles: true,
							composed: true,
						}),
					);
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
		get pressed(): boolean | undefined {
			const value = this.getAttribute('pressed');
			return value === null ? undefined : value !== 'false';
		}
		set pressed(value: boolean | undefined) {
			if (value === undefined) this.removeAttribute('pressed');
			else this.setAttribute('pressed', String(value));
		}
		/** The registered icon to draw; a product subclass may map names onto its glyph set. */
		protected iconName(): string | null {
			return this.getAttribute('icon');
		}
		/** Whether the trailing menu chevron shows. */
		protected showsCaret(): boolean {
			const caret = this.getAttribute('caret');
			return caret !== null && caret !== 'false';
		}
		private sync(): void {
			const label = this.getAttribute('label') ?? '';
			const iconOnly = this.hasAttribute('icon-only');
			this.text.textContent = label;
			// The chevron trails the last line of the label.
			this.text.append(this.caret);
			this.caret.toggleAttribute('hidden', !this.showsCaret());
			this.text.hidden = iconOnly;
			if (iconOnly && label) this.button.setAttribute('aria-label', label);
			else this.button.removeAttribute('aria-label');
			this.button.title = this.getAttribute('title') ?? (iconOnly ? label : '');
			this.button.disabled = this.disabled;
			const shortcuts = this.getAttribute('keyshortcuts');
			if (shortcuts) this.button.setAttribute('aria-keyshortcuts', shortcuts);
			else this.button.removeAttribute('aria-keyshortcuts');
			if (paintIcon(this.svg, this.iconName())) this.svg.setAttribute('data-painted', '');
			else this.svg.removeAttribute('data-painted');
			this.badge.textContent = this.getAttribute('badge') ?? '';
			this.badge.hidden = !this.badge.textContent;
			for (const attr of ['pressed', 'expanded'] as const) {
				const value = this.getAttribute(attr);
				if (value === null) this.button.removeAttribute(`aria-${attr}`);
				else this.button.setAttribute(`aria-${attr}`, value === 'false' ? 'false' : 'true');
			}
		}
	}
	return OfficeUiButton;
});
