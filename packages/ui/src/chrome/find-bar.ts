import { LitElement, html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { live } from 'lit/directives/live.js';
import { OfficeElement, controlStyles, flag } from '../base.js';
import { definer, present } from '../registry.js';
import css from './find-bar.css?raw';

export type OfficeFindInputEvent = CustomEvent<{ query: string }>;
export type OfficeFindStepEvent = CustomEvent<{ direction: 'next' | 'previous' }>;

/**
 * `<office-ui-find-bar>`: Office's on-demand Find bar under the ribbon. Shown while `open`;
 * `show()` opens it and selects the query. Enter steps to the next result, Shift+Enter to the
 * previous; Escape clears a query, or closes the bar when it is already empty. Attributes:
 * `open`, `label`, `input-label`, `placeholder`, `maxlength`, `disabled`, `navigation-disabled`,
 * `previous-label`, `next-label`, `close-label`. Properties: `value`, `status`, `statusTitle`.
 * Events: `office-find-input` `{ query }`, `office-find-step` `{ direction }`, `office-find-close`.
 */
export class OfficeUiFindBar extends OfficeElement {
	static override shadowRootOptions = { ...LitElement.shadowRootOptions, delegatesFocus: true };
	static override styles = controlStyles(css);
	static override properties = {
		open: flag,
		disabled: flag,
		navigationDisabled: { attribute: 'navigation-disabled', ...flag },
		label: { type: String },
		inputLabel: { attribute: 'input-label', type: String },
		placeholder: { type: String },
		maxlength: { type: String },
		previousLabel: { attribute: 'previous-label', type: String },
		nextLabel: { attribute: 'next-label', type: String },
		closeLabel: { attribute: 'close-label', type: String },
		value: { attribute: false },
		status: { attribute: false },
		statusTitle: { attribute: false },
	};
	declare open: boolean;
	declare disabled: boolean;
	declare navigationDisabled: boolean;
	declare label: string | null;
	declare inputLabel: string | null;
	declare placeholder: string | null;
	declare maxlength: string | null;
	declare previousLabel: string | null;
	declare nextLabel: string | null;
	declare closeLabel: string | null;
	declare value: string;
	declare status: string;
	declare statusTitle: string;

	constructor() {
		super();
		this.open = false;
		this.disabled = false;
		this.navigationDisabled = false;
		this.label = null;
		this.inputLabel = null;
		this.placeholder = null;
		this.maxlength = null;
		this.previousLabel = null;
		this.nextLabel = null;
		this.closeLabel = null;
		this.value = '';
		this.status = '';
		this.statusTitle = '';
	}

	private get input(): HTMLInputElement | null {
		return this.renderRoot.querySelector('input');
	}

	/** Open, focus the field and select its query. */
	show(): void {
		this.open = true;
		this.input?.focus();
		this.input?.select();
	}

	/** Close and emit `office-find-close`. */
	close(): void {
		if (!present(this.open)) return;
		this.open = false;
		this.fire('office-find-close', {});
	}

	private step(direction: 'next' | 'previous'): void {
		if (!present(this.navigationDisabled)) this.fire('office-find-step', { direction });
	}

	private onInput(event: Event): void {
		this.value = (event.target as HTMLInputElement).value;
		this.fire('office-find-input', { query: this.value });
	}

	private onKey(event: KeyboardEvent): void {
		if (event.isComposing) return;
		if (event.key === 'Enter') {
			event.preventDefault();
			this.step(event.shiftKey ? 'previous' : 'next');
		} else if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			if (this.value) {
				this.value = '';
				this.fire('office-find-input', { query: '' });
			} else this.close();
		}
	}

	protected override render() {
		const max = Number(this.maxlength);
		const disabled = present(this.disabled);
		const stuck = present(this.navigationDisabled) || disabled;
		const close = this.closeLabel ?? 'Close Find';
		return html`
			<div role="search" aria-label=${this.label ?? 'Find'}>
				<label>
					Find
					<input
						type="search"
						autocomplete="off"
						spellcheck="false"
						aria-describedby="status"
						aria-label=${this.inputLabel ?? 'Find'}
						placeholder=${this.placeholder ?? ''}
						maxlength=${ifDefined(Number.isInteger(max) && max > 0 ? String(max) : undefined)}
						.value=${live(this.value)}
						?disabled=${disabled}
						@input=${this.onInput}
						@keydown=${this.onKey}
					/>
				</label>
				<button
					type="button"
					data-action="previous"
					aria-label=${this.previousLabel ?? 'Previous result'}
					?disabled=${stuck}
					@click=${() => this.step('previous')}
					>Previous</button
				>
				<button
					type="button"
					data-action="next"
					aria-label=${this.nextLabel ?? 'Next result'}
					?disabled=${stuck}
					@click=${() => this.step('next')}
					>Next</button
				>
				<span
					id="status"
					class="status"
					role="status"
					aria-live="polite"
					aria-atomic="true"
					title=${this.statusTitle}
					>${this.status}</span
				>
			</div>
			<button
				type="button"
				class="close"
				data-action="close"
				aria-label=${close}
				title="${close} (Esc)"
				@click=${this.close}
				>×</button
			>
		`;
	}
}

export const defineFindBar = definer('office-ui-find-bar', () => OfficeUiFindBar);
