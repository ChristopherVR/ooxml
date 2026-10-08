import { html } from 'lit';
import { live } from 'lit/directives/live.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import type { OfficeSelectOption } from '../form/select';

export interface FindReplaceFields {
	replacement: string;
	replacementMaxlength: number;
	scope: string;
	scopeOptions: OfficeSelectOption[];
	matchCase: boolean;
	matchCaseDisabled: boolean;
	disabled: boolean;
	replaceDisabled: boolean;
	replaceAllDisabled: boolean;
	options(next: { replacement?: string; scope?: string; matchCase?: boolean }): void;
	replace(mode: 'current' | 'all'): void;
}

/** Optional product-neutral Replace controls; matching, scopes and transactions belong to core. */
export function findReplaceFields(state: FindReplaceFields) {
	return html`<div class="replace-fields">
		<label
			>Replace with
			<textarea
				rows="2"
				aria-label="Replace with"
				spellcheck="false"
				maxlength=${ifDefined(state.replacementMaxlength > 0 ? String(state.replacementMaxlength) : undefined)}
				.value=${live(state.replacement)}
				?disabled=${state.disabled}
				@input=${(event: Event) => state.options({ replacement: (event.target as HTMLTextAreaElement).value })}
			></textarea>
		</label>
		<label
			>Search in
			<office-ui-select
				aria-label="Search scope"
				.value=${state.scope}
				.options=${state.scopeOptions}
				?disabled=${state.disabled}
				@change=${(event: Event) => state.options({ scope: (event.currentTarget as HTMLElement & { value: string }).value })}
			></office-ui-select
		></label>
		<label class="match-case"
			><office-ui-checkbox
				aria-label="Match case"
				.checked=${state.matchCase}
				?disabled=${state.disabled || state.matchCaseDisabled}
				@change=${(event: Event) => state.options({ matchCase: (event.currentTarget as HTMLElement & { checked: boolean }).checked })}
			></office-ui-checkbox
			>Match case</label
		>
		<button
			type="button"
			data-action="replace"
			?disabled=${state.disabled || state.replaceDisabled}
			@click=${() => state.replace('current')}
			>Replace</button
		>
		<button
			type="button"
			data-action="replace-all"
			?disabled=${state.disabled || state.replaceAllDisabled}
			@click=${() => state.replace('all')}
			>Replace All</button
		>
	</div>`;
}
