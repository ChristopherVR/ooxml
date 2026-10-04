import { present } from './registry.js';
import { tok } from './tokens.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CHECK_SVG =
	'<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 8 3.2 3.2L13 4.5"/></svg>';

const CHECKBOX_CSS = `
:host { display: inline-grid; box-sizing: border-box; width: ${tok('--office-checkbox-size')}; height: ${tok('--office-checkbox-size')};
	flex: none; place-items: center; border: ${tok('--office-border-width')} solid ${tok('--office-checkbox-border')};
	border-radius: ${tok('--office-checkbox-radius')}; background: ${tok('--office-checkbox-background')};
	color: ${tok('--office-checkbox-accent-foreground')}; cursor: pointer; vertical-align: middle; }
:host([checked]) { border-color: ${tok('--office-checkbox-accent')};
	background: ${tok('--office-checkbox-accent')}; }
:host(:focus-visible) { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
svg { display: none; width: ${tok('--office-icon-size-sm')}; height: ${tok('--office-icon-size-sm')}; }
:host([checked]) svg { display: block; }
@media (pointer: coarse), (max-width: 767px) {
	:host { width: ${tok('--office-checkbox-size-touch')}; height: ${tok('--office-checkbox-size-touch')}; }
	svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; }
}
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; color: CanvasText; forced-color-adjust: auto; }
	:host([checked]) { border-color: Highlight; background: Highlight; color: HighlightText; }
}
`;

/* Switch metrics are tokens so a host (a title bar) can size it: --office-switch-width, -height,
 * -border-width, -knob-size, -knob-offset, -knob-travel, -track, -track-on, -thumb, -thumb-on. */
const SWITCH_CSS = `
:host { --_w: ${tok('--office-switch-width')}; --_h: ${tok('--office-switch-height')};
	--_b: ${tok('--office-switch-border-width')}; --_k: ${tok('--office-switch-knob-size')};
	--_o: ${tok('--office-switch-knob-offset')};
	display: inline-block; box-sizing: border-box; position: relative; width: var(--_w); height: var(--_h);
	flex: none; border: var(--_b) solid ${tok('--office-switch-border')}; border-radius: ${tok('--office-radius-full')};
	background: ${tok('--office-switch-track')}; cursor: pointer; vertical-align: middle;
	touch-action: manipulation; transition: background-color ${tok('--office-duration')}; }
.knob { display: block; position: absolute; top: calc((var(--_h) - 2 * var(--_b) - var(--_k)) / 2); left: var(--_o);
	width: var(--_k); height: var(--_k); border-radius: 50%; background: ${tok('--office-switch-thumb')};
	box-shadow: ${tok('--office-shadow-sm')}; transition: transform ${tok('--office-duration')}; }
:host([checked]) { border-color: ${tok('--office-switch-track-on')};
	background: ${tok('--office-switch-track-on')}; }
:host([checked]) .knob { transform: translateX(var(--office-switch-knob-travel, calc(var(--_w) - var(--_k) - 2 * var(--_o) - 2 * var(--_b))));
	background: ${tok('--office-switch-thumb-on')}; }
:host(:focus-visible) { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
svg { display: none; }
@media (pointer: coarse), (max-width: 767px) { :host { --_w: ${tok('--office-switch-width-touch')}; --_h: ${tok('--office-switch-height-touch')}; --_k: ${tok('--office-switch-knob-size-touch')}; } }
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; forced-color-adjust: none; }
	.knob { background: CanvasText; }
	:host([checked]) { border-color: Highlight; background: Highlight; }
	:host([checked]) .knob { background: HighlightText; }
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
				// The switch draws a real knob element, so hosts and tests can measure and style it.
				if (role === 'switch') {
					const knob = this.ownerDocument.createElement('span');
					knob.className = 'knob';
					knob.setAttribute('part', 'knob');
					root.append(knob);
				}
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
				this.toggleAttribute('checked', present(value));
			}
			get disabled(): boolean {
				return this.hasAttribute('disabled');
			}
			set disabled(value: boolean) {
				this.toggleAttribute('disabled', present(value));
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
