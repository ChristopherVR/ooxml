import { html } from 'lit';
import { OfficeElement, flag, valueOn } from '../base.js';
import { present } from '../registry.js';

/**
 * Shared implementation of the two-state controls (checkbox, switch). `checked`, `disabled` and
 * `value` attributes and properties; Space (and Enter for the switch) toggles; `input` then
 * `change` bubble from the host; form-associated through ElementInternals where supported;
 * setting a property is silent.
 */
export abstract class OfficeUiCheckable extends OfficeElement {
	static formAssociated = true;
	static override properties = {
		checked: flag,
		disabled: flag,
		value: valueOn,
	};
	declare checked: boolean;
	declare disabled: boolean;
	declare value: string;

	protected abstract readonly semantics: 'checkbox' | 'switch';
	private readonly internals: ElementInternals | undefined;
	private defaultChecked = false;

	constructor() {
		super();
		this.checked = false;
		this.disabled = false;
		this.value = 'on';
		try {
			this.internals = this.attachInternals();
		} catch {
			this.internals = undefined;
		}
		this.addEventListener('click', (event) => {
			if (present(this.disabled)) event.preventDefault();
			else this.toggle();
		});
		this.addEventListener('keydown', (event) => {
			const toggles = event.key === ' ' || (this.semantics === 'switch' && event.key === 'Enter');
			if (!toggles || present(this.disabled)) return;
			event.preventDefault();
			this.toggle();
		});
	}

	override connectedCallback(): void {
		this.defaultChecked = present(this.checked);
		super.connectedCallback();
	}

	formResetCallback(): void {
		this.checked = this.defaultChecked;
	}

	formDisabledCallback(disabled: boolean): void {
		this.disabled = disabled;
	}

	private toggle(): void {
		this.checked = !present(this.checked);
		this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
		this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
	}

	/** The host carries the semantics: role, state, tab stop and form value. */
	protected override willUpdate(): void {
		const checked = present(this.checked);
		const disabled = present(this.disabled);
		this.setAttribute('role', this.semantics);
		this.setAttribute('aria-checked', String(checked));
		this.setAttribute('aria-disabled', String(disabled));
		this.tabIndex = disabled ? -1 : 0;
		this.internals?.setFormValue?.(checked && !disabled ? this.value : null);
	}

	/** The tick, shown through CSS when `checked`. */
	protected renderCheck() {
		return html`<svg
			viewBox="0 0 16 16"
			fill="none"
			stroke="currentColor"
			stroke-width="2"
			stroke-linecap="round"
			stroke-linejoin="round"
			aria-hidden="true"
		>
			<path d="m3 8 3.2 3.2L13 4.5"></path>
		</svg>`;
	}
}
