import { LitElement, html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { live } from 'lit/directives/live.js';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, present } from '../registry';
import css from './find-bar.css?raw';
import { findReplaceFields } from './find-replace-fields';
import type { OfficeSelectOption } from '../form/select';

export type OfficeFindInputEvent = CustomEvent<{ query: string }>;
export type OfficeFindStepEvent = CustomEvent<{ direction: 'next' | 'previous' }>;
export type OfficeFindOptionsEvent = CustomEvent<{
	replacement: string;
	scope: string;
	matchCase: boolean;
}>;
export type OfficeFindReplaceEvent = CustomEvent<{ mode: 'current' | 'all' }>;

/**
 * `<office-ui-find-bar>`: Office's on-demand Find bar under the ribbon. Shown while `open`;
 * `show()` opens it and selects the query. Enter steps to the next result, Shift+Enter to the
 * previous; Escape clears a query, or closes the bar when it is already empty. Attributes:
 * `open`, `label`, `input-label`, `placeholder`, `maxlength`, `disabled`, `navigation-disabled`,
 * `previous-label`, `next-label`, `close-label`. Properties: `value`, `status`, `statusTitle`.
 * Events: `office-find-input` `{ query }`, `office-find-step` `{ direction }`, `office-find-close`.
 * Opt-in `replace-mode` adds a multiline `replacement`, `scope`/`scopeOptions`, `matchCase`,
 * `match-case-disabled`, `replace-disabled`, `replace-all-disabled`, `replacement-maxlength`
 * and inert `error` text. `office-find-options` reports replacement/scope/matchCase;
 * `office-find-replace` requests current/all. The host owns matching and edit transactions.
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
		replaceMode: { attribute: 'replace-mode', ...flag },
		replacement: { attribute: false },
		replacementMaxlength: { attribute: 'replacement-maxlength', type: Number },
		scope: { attribute: false },
		scopeOptions: { attribute: false },
		matchCase: { attribute: 'match-case', ...flag },
		matchCaseDisabled: { attribute: 'match-case-disabled', ...flag },
		replaceDisabled: { attribute: 'replace-disabled', ...flag },
		replaceAllDisabled: { attribute: 'replace-all-disabled', ...flag },
		error: { attribute: false },
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
	declare replaceMode: boolean;
	declare replacement: string;
	declare replacementMaxlength: number;
	declare scope: string;
	declare scopeOptions: OfficeSelectOption[];
	declare matchCase: boolean;
	declare matchCaseDisabled: boolean;
	declare replaceDisabled: boolean;
	declare replaceAllDisabled: boolean;
	declare error: string;

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
		this.replaceMode = false;
		this.replacement = '';
		this.replacementMaxlength = 0;
		this.scope = '';
		this.scopeOptions = [];
		this.matchCase = true;
		this.matchCaseDisabled = false;
		this.replaceDisabled = false;
		this.replaceAllDisabled = false;
		this.error = '';
		this.addEventListener('keydown', (event) => {
			if (present(this.replaceMode) && event.key === 'Escape' && !event.defaultPrevented)
				this.onKey(event);
		});
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

	private replaceOptions(next: {
		replacement?: string;
		scope?: string;
		matchCase?: boolean;
	}): void {
		if (present(this.disabled)) return;
		if (next.replacement !== undefined) this.replacement = next.replacement;
		if (next.scope !== undefined) this.scope = next.scope;
		if (next.matchCase !== undefined && !present(this.matchCaseDisabled))
			this.matchCase = next.matchCase;
		this.fire('office-find-options', {
			replacement: this.replacement,
			scope: this.scope,
			matchCase: this.matchCase,
		});
	}

	protected override render() {
		const max = Number(this.maxlength);
		const disabled = present(this.disabled);
		const stuck = present(this.navigationDisabled) || disabled;
		const close = this.closeLabel ?? 'Close Find';
		return html`
			<div class="find-content">
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
				${
					present(this.replaceMode)
						? findReplaceFields({
								replacement: this.replacement,
								replacementMaxlength: this.replacementMaxlength,
								scope: this.scope,
								scopeOptions: this.scopeOptions,
								matchCase: this.matchCase,
								matchCaseDisabled: present(this.matchCaseDisabled),
								disabled,
								replaceDisabled: present(this.replaceDisabled),
								replaceAllDisabled: present(this.replaceAllDisabled),
								options: (next) => this.replaceOptions(next),
								replace: (mode) => {
									if (
										!disabled &&
										!present(mode === 'current' ? this.replaceDisabled : this.replaceAllDisabled)
									)
										this.fire('office-find-replace', { mode });
								},
							})
						: null
				}
				${this.error ? html`<p class="find-error" role="alert">${this.error}</p>` : null}
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
