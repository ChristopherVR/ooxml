import { LitElement, html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer, present } from '../registry';
import css from './menu-item.css?raw';

/**
 * One menu entry: `command`, `label`, `icon`, `disabled`, `checked` (renders a
 * `menuitemcheckbox`), `title`. Activation emits `office-command` `{ command }`.
 */
export class OfficeUiMenuItem extends OfficeElement {
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		label: { type: String },
		icon: { type: String },
		command: { type: String },
		checked: { type: String },
		disabled: flag,
	};
	declare label: string;
	declare icon: string | null;
	declare command: string | null;
	declare checked: string | null;
	declare disabled: boolean;

	/** `title` is a native attribute, so it is watched rather than declared as a property. */
	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'title'];
	}

	constructor() {
		super();
		this.label = '';
		this.icon = null;
		this.command = null;
		this.checked = null;
		this.disabled = false;
	}

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		super.attributeChangedCallback(name, old, value);
		if (name === 'title') this.requestUpdate();
	}

	/** Focus the menuitem itself, so arrow navigation never depends on delegatesFocus. */
	override focus(options?: FocusOptions): void {
		this.renderRoot.querySelector('button')?.focus(options);
	}

	private onClick(): void {
		if (!present(this.disabled) && this.command)
			this.fire('office-command', { command: this.command });
	}

	protected override render() {
		const checkable = this.checked !== null;
		const on = checkable && this.checked !== 'false';
		return html`
			<button
				type="button"
				tabindex="-1"
				role=${checkable ? 'menuitemcheckbox' : 'menuitem'}
				aria-checked=${ifDefined(checkable ? String(on) : undefined)}
				title=${this.getAttribute('title') ?? ''}
				?disabled=${present(this.disabled)}
				@click=${this.onClick}
			>
				<span class="check" aria-hidden="true">${on ? '✓' : ''}</span>${glyph(this.icon)}<span
					>${this.label}</span
				>
			</button>
		`;
	}
}

export const defineMenuItem = definer('office-ui-menu-item', () => OfficeUiMenuItem);
