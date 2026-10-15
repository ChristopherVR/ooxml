import { html, type TemplateResult } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';

/**
 * The control rows of the shared File > Options dialog: toggle, select, number and text, each a
 * `label` row (label start, control end) in the PowerPoint arrangement. Pure templates over a
 * draft record, so the dialog element stays about state and layout.
 */
export type OfficeOptionValue = boolean | number | string;
export type OfficeOptionValues = Record<string, OfficeOptionValue>;

export interface OfficeOptionChoice {
	value: string;
	label: string;
}

interface OptionBase {
	/** Key into the values record. */
	key: string;
	label: string;
	/** Optional "(i)" tooltip, as Office's info bubbles. */
	info?: string;
	/** Renders indented under the preceding control. */
	indent?: boolean;
	/** Shown disabled with this reason, for settings the product cannot honour yet. */
	disabled?: string;
	/** Enabled only while the toggle with this key is on (Excel's iterative calculation fields). */
	enabledBy?: string;
}

export type OfficeOptionControl =
	| (OptionBase & { kind: 'toggle' })
	| (OptionBase & { kind: 'select'; choices: readonly OfficeOptionChoice[] })
	| (OptionBase & { kind: 'number'; min: number; max: number; step?: number; unit?: string })
	| (OptionBase & { kind: 'text'; maxLength?: number });

/** Clamp a number edit into range; `undefined` when it is not a finite number. */
export function clampOptionNumber(raw: string, min: number, max: number): number | undefined {
	const parsed = Number(raw);
	return Number.isFinite(parsed) && raw.trim() !== ''
		? Math.min(max, Math.max(min, parsed))
		: undefined;
}

export interface OptionRowHost {
	readonly draft: OfficeOptionValues;
	set(key: string, value: OfficeOptionValue): void;
}

export function optionRow(host: OptionRowHost, control: OfficeOptionControl): TemplateResult {
	const { draft } = host;
	const value = draft[control.key];
	const locked =
		Boolean(control.disabled) || (control.enabledBy && draft[control.enabledBy] !== true);
	const disabled = Boolean(locked);
	const info = control.info
		? html`<span class="info" title=${control.info} role="img" aria-label=${control.info}>i</span>`
		: '';
	const label = html`<span class="label">${control.label}${info}</span>`;
	let body: TemplateResult;
	if (control.kind === 'toggle')
		body = html`${label}<input
				type="checkbox"
				.checked=${value === true}
				?disabled=${disabled}
				@change=${(event: Event) => host.set(control.key, (event.target as HTMLInputElement).checked)}
			/>`;
	else if (control.kind === 'select') {
		const current = String(value ?? control.choices[0]?.value ?? '');
		body = html`${label}<select
				aria-label=${control.label}
				?disabled=${disabled}
				@change=${(event: Event) => host.set(control.key, (event.target as HTMLSelectElement).value)}
			>
				${control.choices.map(
					(choice) =>
						html`<option value=${choice.value} .selected=${choice.value === current}
							>${choice.label}</option
						>`,
				)}
			</select>`;
	} else if (control.kind === 'number')
		body = html`${label}<span class="number"
				><input
					type="number"
					min=${control.min}
					max=${control.max}
					step=${ifDefined(control.step ? String(control.step) : undefined)}
					aria-label=${control.label}
					.value=${value === undefined ? '' : String(value)}
					?disabled=${disabled}
					@change=${(event: Event) => {
						const field = event.target as HTMLInputElement;
						const next = clampOptionNumber(field.value, control.min, control.max);
						if (next === undefined) field.value = String(draft[control.key] ?? '');
						else {
							field.value = String(next);
							host.set(control.key, next);
						}
					}}
				/>${control.unit ? html`<span class="unit">${control.unit}</span>` : ''}</span
			>`;
	else
		body = html`${label}<input
				type="text"
				maxlength=${ifDefined(control.maxLength ? String(control.maxLength) : undefined)}
				aria-label=${control.label}
				.value=${value === undefined ? '' : String(value)}
				?disabled=${disabled}
				@input=${(event: Event) => host.set(control.key, (event.target as HTMLInputElement).value)}
			/>`;
	return html`<label
		class="row${control.indent ? ' indent' : ''}"
		data-key=${control.key}
		?data-disabled=${disabled}
		title=${ifDefined(control.disabled)}
		>${body}</label
	>`;
}
