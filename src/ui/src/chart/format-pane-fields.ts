/**
 * Templates for the fields of the chart Format pane, drawn from their
 * descriptors (moved from PowerPoint's imperative
 * `web-components/chart-format-pane-fields.ts`). Choice controls use the tags
 * the element names, so a product subclass keeps its own (`pptx-ui-radio`...).
 * Values of boxes and selects are written by the element after each render
 * (`data-value`), skipping the control being typed in.
 *
 * @module chart/format-pane-fields
 */
import { html, nothing, type TemplateResult } from 'lit';
import { live } from 'lit/directives/live.js';
import { html as staticHtml, unsafeStatic } from 'lit/static-html.js';

import { rangeControl } from '../form/range-control';
import type { ChartFormatField, ChartFormatFieldValue, ChartFormatText } from './format-pane-types';

export interface ChartFormatControlTags {
	select: string;
	radio: string;
	checkbox: string;
}

export interface FieldContext {
	tags: ChartFormatControlTags;
	tr(text: ChartFormatText): string;
	change(id: string, value: ChartFormatFieldValue): void;
	/** Shows `value` in the box of field `id` without committing it. */
	preview(id: string, value: number, text: string): void;
	locked: boolean;
	autoLabel: string;
}

interface Checkable extends HTMLElement {
	checked?: boolean;
	value: string;
}

const caption = (text: string) => html`<span class="caption">${text}</span>`;

function readValue(field: ChartFormatField, input: Checkable): ChartFormatFieldValue {
	if (field.kind === 'checkbox') return Boolean(input.checked);
	if (field.kind === 'number') return input.value.trim() === '' ? '' : Number(input.value);
	return input.value;
}

function radioGroup(
	field: Extract<ChartFormatField, { kind: 'radio' }>,
	ctx: FieldContext,
): TemplateResult {
	const tag = unsafeStatic(ctx.tags.radio);
	return html`<fieldset class="radio-group" data-chart-format-field=${field.id}
		><legend>${ctx.tr(field.label)}</legend>${field.choices.map(
			(choice) =>
				html`<label class="choice" title=${choice.hint ? ctx.tr(choice.hint) : nothing}
					>${staticHtml`<${tag}
					name=${`cfp-${field.id}`}
					.value=${choice.value}
					data-chart-format-choice=${choice.value}
					.checked=${live(choice.value === field.value)}
					?disabled=${ctx.locked || Boolean(choice.disabled) || Boolean(field.disabled)}
					@change=${(event: Event) => {
						if ((event.currentTarget as Checkable).checked) ctx.change(field.id, choice.value);
					}}
				></${tag}>`}${caption(ctx.tr(choice.label))}</label
				>`,
		)}</fieldset
	>`;
}

function fileField(
	field: Extract<ChartFormatField, { kind: 'file' }>,
	ctx: FieldContext,
): TemplateResult {
	const pick = (event: Event) => {
		const input = event.currentTarget as HTMLInputElement;
		const file = input.files?.[0];
		if (!file) return;
		const reader = new FileReader();
		reader.addEventListener('load', () => {
			if (typeof reader.result === 'string') ctx.change(field.id, reader.result);
		});
		reader.readAsDataURL(file);
		input.value = '';
	};
	return html`<div
		class="field field-file"
		data-chart-format-field=${field.id}
		data-chart-format-file-set=${String(field.value)}
		>${caption(ctx.tr(field.label))}<button
			type="button"
			class="file-button"
			?disabled=${ctx.locked || Boolean(field.disabled)}
			@click=${(event: Event) =>
				(
					(event.currentTarget as HTMLElement).nextElementSibling as HTMLInputElement | null
				)?.click()}
			>${ctx.tr(field.button)}</button
		><input type="file" accept=${field.accept} hidden @change=${pick}
	/></div>`;
}

function control(field: ChartFormatField, ctx: FieldContext, label: string): TemplateResult {
	const disabled = ctx.locked || Boolean(field.disabled);
	const commit = (event: Event) =>
		ctx.change(field.id, readValue(field, event.currentTarget as Checkable));
	if (field.kind === 'select') {
		const tag = unsafeStatic(ctx.tags.select);
		return staticHtml`<${tag}
			aria-label=${label}
			data-value=${field.value}
			?disabled=${disabled}
			@change=${commit}
		>${field.choices.map(
			(choice) =>
				html`<option
					value=${choice.value}
					.selected=${live(choice.value === field.value)}
					?disabled=${Boolean(choice.disabled)}
					>${ctx.tr(choice.label)}</option
				>`,
		)}</${tag}>`;
	}
	if (field.kind === 'checkbox') {
		const tag = unsafeStatic(ctx.tags.checkbox);
		return staticHtml`<${tag}
			aria-label=${label}
			.checked=${live(field.value)}
			?disabled=${disabled}
			@change=${commit}
		></${tag}>`;
	}
	const number = field.kind === 'number' ? field : undefined;
	return html`<input
		type=${field.kind === 'color' ? 'color' : number ? 'number' : 'text'}
		aria-label=${label}
		data-value=${String(field.value)}
		min=${number?.min ?? nothing}
		max=${number?.max ?? nothing}
		step=${number ? (number.step ?? 'any') : nothing}
		placeholder=${number?.auto ? ctx.autoLabel : nothing}
		inputmode=${number ? 'decimal' : nothing}
		?disabled=${disabled}
		@change=${commit}
	/>`;
}

function slider(
	field: Extract<ChartFormatField, { kind: 'number' }>,
	ctx: FieldContext,
	label: string,
): TemplateResult {
	const value = field.value === '' ? (field.min ?? 0) : field.value;
	let last = value;
	return html`<div class="slider" data-chart-format-slider=${field.id}
		>${rangeControl({
			label,
			value,
			valueText: `${value}${field.unit ?? ''}`,
			min: field.min ?? 0,
			max: field.max ?? 100,
			...(field.step === undefined ? {} : { step: field.step }),
			disabled: ctx.locked || Boolean(field.disabled),
			onInput: (next) => {
				// The box previews the number while dragging; the slider's change commits it.
				last = next;
				ctx.preview(field.id, next, `${next}${field.unit ?? ''}`);
			},
			onChange: () => ctx.change(field.id, last),
		})}</div
	>`;
}

/** The template of one field. */
export function fieldTemplate(field: ChartFormatField, ctx: FieldContext): TemplateResult {
	if (field.kind === 'radio') return radioGroup(field, ctx);
	if (field.kind === 'file') return fileField(field, ctx);
	const label = ctx.tr(field.label);
	const title = field.hint ? ctx.tr(field.hint) : nothing;
	if (field.kind === 'checkbox')
		return html`<label
			class="field field-checkbox"
			data-chart-format-field=${field.id}
			title=${title}
			>${control(field, ctx, label)}${caption(label)}</label
		>`;
	const unit =
		field.kind === 'number' && field.unit ? html`<span class="unit">${field.unit}</span>` : nothing;
	return html`<label
			class="field field-${field.kind}"
			data-chart-format-field=${field.id}
			title=${title}
			>${caption(label)}${control(field, ctx, label)}${unit}</label
		>${field.kind === 'number' && field.slider ? slider(field, ctx, label) : nothing}`;
}
