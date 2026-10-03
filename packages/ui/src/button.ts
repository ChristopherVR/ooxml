import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeCommandEvent = CustomEvent<{ command: string }>;

const CSS = `
:host { display: inline-flex; vertical-align: middle; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center;
	gap: ${tok('--office-space-1-5')}; min-width: ${tok('--office-target-size')}; min-height: ${tok('--office-target-size')};
	padding: 0 ${tok('--office-space-2')}; border: ${tok('--office-border-width')} solid transparent; border-radius: ${tok('--office-radius')};
	background: transparent; color: ${tok('--office-foreground')}; font: inherit; font-size: ${tok('--office-font-size-sm')};
	cursor: pointer; white-space: nowrap; }
:host([variant="stacked"]) button { flex-direction: column; padding: ${tok('--office-space-1')} ${tok('--office-space-2')}; }
button:hover:not(:disabled) { background: ${tok('--office-surface')}; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
button[aria-pressed="true"], button[aria-expanded="true"] { border-color: ${tok('--office-accent')};
	background: ${tok('--office-surface')}; }
button:disabled { opacity: .5; cursor: not-allowed; }
svg { width: ${tok('--office-icon-size-ml')}; height: ${tok('--office-icon-size-ml')}; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
	stroke-linecap: round; stroke-linejoin: round; }
svg:not([data-painted]) { display: none; }
@media (forced-colors: active) {
	button { border-color: ButtonText; color: ButtonText; background: ButtonFace; }
	button:disabled { color: GrayText; border-color: GrayText; }
	button[aria-pressed="true"], button[aria-expanded="true"] { background: Highlight; color: HighlightText; }
}
`;

/**
 * Command button. Attributes: `label`, `icon`, `command`, `disabled`, `pressed` (aria-pressed),
 * `expanded` (aria-expanded), `icon-only`, `variant="stacked"`, `title`, `keyshortcuts`
 * (forwarded as aria-keyshortcuts, e.g. `Control+Z`).
 * Activation (pointer, Enter, Space) emits one bubbling, composed `office-command`
 * event with `{ command }`; the host decides what it does. Nothing happens when disabled.
 */
export const defineButton = definer('office-ui-button', () => {
	class OfficeUiButton extends HTMLElement {
		static observedAttributes = [
			'label',
			'icon',
			'command',
			'disabled',
			'pressed',
			'expanded',
			'icon-only',
			'title',
			'keyshortcuts',
		];
		private readonly button: HTMLButtonElement;
		private readonly svg: SVGSVGElement;
		private readonly text: HTMLSpanElement;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			this.button = this.ownerDocument.createElement('button');
			this.button.type = 'button';
			this.button.setAttribute('part', 'button');
			this.svg = createIconSvg(this.ownerDocument);
			this.text = this.ownerDocument.createElement('span');
			this.button.append(this.svg, this.text);
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
		get pressed(): boolean | undefined {
			const value = this.getAttribute('pressed');
			return value === null ? undefined : value !== 'false';
		}
		set pressed(value: boolean | undefined) {
			if (value === undefined) this.removeAttribute('pressed');
			else this.setAttribute('pressed', String(value));
		}
		private sync(): void {
			const label = this.getAttribute('label') ?? '';
			const iconOnly = this.hasAttribute('icon-only');
			this.text.textContent = label;
			this.text.hidden = iconOnly;
			if (iconOnly && label) this.button.setAttribute('aria-label', label);
			else this.button.removeAttribute('aria-label');
			this.button.title = this.getAttribute('title') ?? (iconOnly ? label : '');
			this.button.disabled = this.disabled;
			const shortcuts = this.getAttribute('keyshortcuts');
			if (shortcuts) this.button.setAttribute('aria-keyshortcuts', shortcuts);
			else this.button.removeAttribute('aria-keyshortcuts');
			if (paintIcon(this.svg, this.getAttribute('icon'))) this.svg.setAttribute('data-painted', '');
			else this.svg.removeAttribute('data-painted');
			for (const attr of ['pressed', 'expanded'] as const) {
				const value = this.getAttribute(attr);
				if (value === null) this.button.removeAttribute(`aria-${attr}`);
				else this.button.setAttribute(`aria-${attr}`, value === 'false' ? 'false' : 'true');
			}
		}
	}
	return OfficeUiButton;
});
