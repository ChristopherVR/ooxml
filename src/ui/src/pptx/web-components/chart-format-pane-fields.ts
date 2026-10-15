/**
 * DOM for one field of the chart Format pane, built from its descriptor and
 * repainted in place. Choices use the shared `pptx-ui-radio`,
 * `pptx-ui-checkbox` and `pptx-ui-select` controls, as every pptx panel does.
 *
 * @module web-components/chart-format-pane-fields
 */
import type { ChartFormatField, ChartFormatFieldValue, ChartFormatText } from '../render';

export type PaneTranslate = (text: ChartFormatText) => string;

/** A built field: its root node and a value repaint that skips the focused control. */
export interface PaneFieldView {
	node: HTMLElement;
	paint(field: ChartFormatField, locked: boolean): void;
}

/** The value/checked/disabled surface the shared form controls and native inputs share. */
interface ControlLike extends HTMLElement {
	value: string;
	checked?: boolean;
	disabled: boolean;
	name?: string;
}

let fieldSerial = 0;

function caption(doc: Document, text: string): HTMLSpanElement {
	const span = doc.createElement('span');
	span.className = 'caption';
	span.textContent = text;
	return span;
}

function readValue(field: ChartFormatField, input: ControlLike): ChartFormatFieldValue {
	if (field.kind === 'checkbox') {
		return Boolean(input.checked);
	}
	if (field.kind === 'number') {
		return input.value.trim() === '' ? '' : Number(input.value);
	}
	return input.value;
}

function radioGroup(
	doc: Document,
	field: Extract<ChartFormatField, { kind: 'radio' }>,
	tr: PaneTranslate,
	change: (id: string, value: ChartFormatFieldValue) => void,
): PaneFieldView {
	const name = `cfp-${++fieldSerial}`;
	const set = doc.createElement('fieldset');
	set.className = 'radio-group';
	set.dataset.chartFormatField = field.id;
	const legend = doc.createElement('legend');
	legend.textContent = tr(field.label);
	set.append(legend);
	const radios = field.choices.map((choice) => {
		const label = doc.createElement('label');
		label.className = 'choice';
		const radio = doc.createElement('pptx-ui-radio') as ControlLike;
		radio.name = name;
		radio.value = choice.value;
		radio.dataset.chartFormatChoice = choice.value;
		radio.addEventListener('change', () => {
			if (radio.checked) {
				change(field.id, choice.value);
			}
		});
		label.append(radio, caption(doc, tr(choice.label)));
		if (choice.hint) {
			label.title = tr(choice.hint);
		}
		set.append(label);
		return { radio, choice };
	});
	return {
		node: set,
		paint(next, locked) {
			if (next.kind !== 'radio') {
				return;
			}
			for (const { radio, choice } of radios) {
				radio.checked = choice.value === next.value;
				radio.disabled = locked || Boolean(choice.disabled) || Boolean(next.disabled);
			}
		},
	};
}

function control(doc: Document, field: ChartFormatField, tr: PaneTranslate, autoLabel: string) {
	if (field.kind === 'select') {
		const select = doc.createElement('pptx-ui-select') as ControlLike;
		select.setAttribute('aria-label', tr(field.label));
		for (const choice of field.choices) {
			const option = doc.createElement('option');
			option.value = choice.value;
			option.textContent = tr(choice.label);
			option.disabled = Boolean(choice.disabled);
			select.append(option);
		}
		return select;
	}
	if (field.kind === 'checkbox') {
		return doc.createElement('pptx-ui-checkbox') as ControlLike;
	}
	const box = doc.createElement('input');
	box.type = field.kind === 'color' ? 'color' : field.kind === 'number' ? 'number' : 'text';
	if (field.kind === 'number') {
		if (field.min !== undefined) box.min = String(field.min);
		if (field.max !== undefined) box.max = String(field.max);
		box.step = field.step !== undefined ? String(field.step) : 'any';
		if (field.auto) box.placeholder = autoLabel;
		box.inputMode = 'decimal';
	}
	return box as ControlLike;
}

/** Build the control for `field`; `change` reports a committed value. */
export function createPaneField(
	doc: Document,
	field: ChartFormatField,
	tr: PaneTranslate,
	change: (id: string, value: ChartFormatFieldValue) => void,
	autoLabel: string,
): PaneFieldView {
	if (field.kind === 'radio') {
		return radioGroup(doc, field, tr, change);
	}
	const label = doc.createElement('label');
	label.className = `field field-${field.kind}`;
	label.dataset.chartFormatField = field.id;
	const input = control(doc, field, tr, autoLabel);
	input.addEventListener('change', () => change(field.id, readValue(field, input)));
	if (field.kind === 'checkbox') {
		label.append(input, caption(doc, tr(field.label)));
	} else {
		label.append(caption(doc, tr(field.label)), input);
		if (field.kind === 'number' && field.unit) {
			const unit = doc.createElement('span');
			unit.className = 'unit';
			unit.textContent = field.unit;
			label.append(unit);
		}
	}
	if (field.hint) {
		label.title = tr(field.hint);
	}
	return {
		node: label,
		paint(next, locked) {
			input.disabled = locked || Boolean(next.disabled);
			if (input.matches(':focus') || input.matches(':focus-within')) {
				return;
			}
			if (next.kind === 'checkbox') {
				input.checked = next.value;
			} else if (next.kind !== 'radio') {
				input.value = String(next.value);
			}
		},
	};
}
