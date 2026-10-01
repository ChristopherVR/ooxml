import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeCommandEvent = CustomEvent<{ command: string }>;

const CSS = `
:host { display: inline-flex; vertical-align: middle; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center;
	gap: 6px; min-width: var(--office-target-size, 28px); min-height: var(--office-target-size, 28px);
	padding: 0 8px; border: 1px solid transparent; border-radius: var(--office-radius, 4px);
	background: transparent; color: var(--office-foreground, #1f2937); font: inherit; font-size: 12px;
	cursor: pointer; white-space: nowrap; }
:host([variant="stacked"]) button { flex-direction: column; padding: 4px 8px; }
button:hover:not(:disabled) { background: var(--office-surface, #f3f4f6); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
button[aria-pressed="true"], button[aria-expanded="true"] { border-color: var(--office-accent, #2563eb);
	background: var(--office-surface, #f3f4f6); }
button:disabled { opacity: .5; cursor: not-allowed; }
svg { width: 18px; height: 18px; fill: none; stroke: currentColor; stroke-width: 1.5;
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
 * `expanded` (aria-expanded), `icon-only`, `variant="stacked"`, `title`.
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
