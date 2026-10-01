import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CHECK_SVG =
	'<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 8 3.2 3.2L13 4.5"/></svg>';

const CHECKBOX_CSS = `
:host { display: inline-grid; box-sizing: border-box; width: 16px; height: 16px; flex: none;
	place-items: center; border: 1px solid var(--office-border, #d1d5db); border-radius: 3px;
	background: var(--office-background, #fff); color: var(--office-accent-foreground, #fff);
	cursor: pointer; vertical-align: middle; }
:host([checked]) { border-color: var(--office-accent, #2563eb); background: var(--office-accent, #2563eb); }
:host(:focus-visible) { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 2px; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
svg { display: none; width: 12px; height: 12px; }
:host([checked]) svg { display: block; }
@media (pointer: coarse), (max-width: 767px) { :host { width: 22px; height: 22px; } svg { width: 16px; height: 16px; } }
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; color: CanvasText; forced-color-adjust: auto; }
	:host([checked]) { border-color: Highlight; background: Highlight; color: HighlightText; }
}
`;

const SWITCH_CSS = `
:host { display: inline-block; box-sizing: border-box; position: relative; width: 32px; height: 18px;
	flex: none; border: 1px solid var(--office-border, #d1d5db); border-radius: 9px;
	background: var(--office-surface, #f3f4f6); cursor: pointer; vertical-align: middle; }
:host::after { content: ""; position: absolute; top: 2px; left: 2px; width: 12px; height: 12px;
	border-radius: 50%; background: var(--office-foreground, #1f2937); transition: transform .12s; }
:host([checked]) { border-color: var(--office-accent, #2563eb); background: var(--office-accent, #2563eb); }
:host([checked])::after { transform: translateX(14px); background: var(--office-accent-foreground, #fff); }
:host(:focus-visible) { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 2px; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
svg { display: none; }
@media (pointer: coarse), (max-width: 767px) { :host { min-height: 24px; width: 44px; height: 24px; border-radius: 12px; } }
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; forced-color-adjust: none; }
	:host::after { background: CanvasText; }
	:host([checked]) { border-color: Highlight; background: Highlight; }
	:host([checked])::after { background: HighlightText; }
}
`;

/**
 * Shared implementation of the two-state controls. `checked`, `disabled`, `value` attributes
 * and properties; Space (and Enter for the switch) toggles; `input` then `change` bubble from the
 * host; form-associated through ElementInternals where supported; setting a property is silent.
 */
function makeCheckable(role: 'checkbox' | 'switch', css: string): () => CustomElementConstructor {
	return () =>
		class OfficeUiCheckable extends HTMLElement {
			static formAssociated = true;
			static observedAttributes = ['checked', 'disabled', 'value'];
			private readonly internals: ElementInternals | undefined;
			private defaultChecked = false;
			constructor() {
				super();
				try {
					this.internals = this.attachInternals();
				} catch {
					this.internals = undefined;
				}
				const root = this.attachShadow({ mode: 'open' });
				attachStyles(root, controlCss(css));
				const template = this.ownerDocument.createElement('template');
				template.innerHTML = CHECK_SVG;
				root.append(template.content.cloneNode(true));
				this.addEventListener('click', (event) => {
					if (this.disabled) event.preventDefault();
					else this.toggle();
				});
				this.addEventListener('keydown', (event) => {
					const key = event.key;
					if ((key !== ' ' && !(role === 'switch' && key === 'Enter')) || this.disabled) return;
					event.preventDefault();
					this.toggle();
				});
			}
			connectedCallback(): void {
				this.defaultChecked = this.checked;
				this.sync();
			}
			attributeChangedCallback(): void {
				this.sync();
			}
			get checked(): boolean {
				return this.hasAttribute('checked');
			}
			set checked(value: boolean) {
				this.toggleAttribute('checked', Boolean(value));
			}
			get disabled(): boolean {
				return this.hasAttribute('disabled');
			}
			set disabled(value: boolean) {
				this.toggleAttribute('disabled', Boolean(value));
			}
			get value(): string {
				return this.getAttribute('value') ?? 'on';
			}
			set value(value: string) {
				this.setAttribute('value', String(value));
			}
			formResetCallback(): void {
				this.checked = this.defaultChecked;
			}
			formDisabledCallback(disabled: boolean): void {
				this.disabled = disabled;
			}
			private toggle(): void {
				this.checked = !this.checked;
				this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
			}
			private sync(): void {
				this.setAttribute('role', role);
				this.setAttribute('aria-checked', String(this.checked));
				this.setAttribute('aria-disabled', String(this.disabled));
				this.tabIndex = this.disabled ? -1 : 0;
				this.internals?.setFormValue?.(this.checked && !this.disabled ? this.value : null);
			}
		};
}

/** `<office-ui-checkbox checked disabled value aria-label>`. */
export const defineCheckbox = definer(
	'office-ui-checkbox',
	makeCheckable('checkbox', CHECKBOX_CSS),
);
/** `<office-ui-switch>`: same contract as the checkbox with `role="switch"`. */
export const defineSwitch = definer('office-ui-switch', makeCheckable('switch', SWITCH_CSS));
