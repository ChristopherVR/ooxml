import { LitElement, html } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer } from '../registry';
import css from './status-item.css?raw';

/**
 * One status entry: `label` and `value` attributes (textual, so assistive tech reads both).
 * With the `interactive` attribute it renders a button and emits `office-status-activate`
 * `{ id }` (`id` attribute) on click, Enter or Space; zoom or language pickers are examples.
 */
export class OfficeUiStatusItem extends OfficeElement {
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		label: { type: String },
		value: { type: String, reflect: true },
		interactive: flag,
	};
	declare label: string;
	declare value: string;
	declare interactive: boolean;

	constructor() {
		super();
		this.label = '';
		this.value = '';
		this.interactive = false;
	}

	protected override render() {
		const content = html`${this.label ? html`<span class="label">${this.label}</span>` : ''}<span
				class="value"
				>${this.value}</span
			>`;
		return this.interactive
			? html`<button
					type="button"
					class="wrap"
					@click=${() => this.fire('office-status-activate', { id: this.getAttribute('id') ?? '' })}
					>${content}</button
				>`
			: html`<span class="wrap">${content}</span>`;
	}
}

export const defineStatusItem = definer('office-ui-status-item', () => OfficeUiStatusItem);
