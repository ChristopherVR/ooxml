import { LitElement, html } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, present } from '../registry';
import css from './search-field.css?raw';

/**
 * A single styled search input (title bar, recent files, panel filters). Moved from pptx-viewer
 * `pptx-ui-search`. `value`, `placeholder`, `disabled`, `aria-label` (defaults to the
 * placeholder) and `variant="titlebar"` (compact). `input` and `change` bubble from the host;
 * setting `value` is silent. Tokens: `--office-field-*`.
 */
export class OfficeUiSearch extends OfficeElement {
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		value: { type: String },
		placeholder: { type: String, reflect: true },
		disabled: flag,
		variant: { type: String, reflect: true },
	};
	declare value: string;
	declare placeholder: string | null;
	declare disabled: boolean;
	declare variant: string;

	/** `aria-label` names the inner input, so it is watched rather than declared as a property. */
	static override get observedAttributes(): string[] {
		return [...super.observedAttributes, 'aria-label'];
	}

	constructor() {
		super();
		this.value = '';
		this.placeholder = null;
		this.disabled = false;
		this.variant = '';
	}

	override attributeChangedCallback(name: string, old: string | null, value: string | null): void {
		super.attributeChangedCallback(name, old, value);
		if (name === 'aria-label') this.requestUpdate();
	}

	private get input(): HTMLInputElement | null {
		return this.renderRoot.querySelector('input');
	}

	override focus(options?: FocusOptions): void {
		this.input?.focus(options);
	}

	select(): void {
		this.input?.select();
	}

	private onInput(event: Event): void {
		event.stopPropagation();
		this.value = (event.target as HTMLInputElement).value;
		this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
	}

	private onChange(event: Event): void {
		event.stopPropagation();
		this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
	}

	protected override render() {
		return html`
			<svg
				part="icon"
				viewBox="0 0 24 24"
				fill="none"
				stroke="currentColor"
				stroke-width="2"
				stroke-linecap="round"
				aria-hidden="true"
			>
				<circle cx="11" cy="11" r="8"></circle>
				<path d="m21 21-4.35-4.35"></path>
			</svg>
			<input
				part="input"
				type="search"
				placeholder=${this.placeholder ?? ''}
				aria-label=${this.getAttribute('aria-label') ?? this.placeholder ?? ''}
				.value=${this.value}
				?disabled=${present(this.disabled)}
				@input=${this.onInput}
				@change=${this.onChange}
			/>
		`;
	}
}

export const defineSearchField = definer('office-ui-search', () => OfficeUiSearch);
