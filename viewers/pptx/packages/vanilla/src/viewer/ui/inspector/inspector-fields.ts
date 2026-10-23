import {
	createLiveColorPatchInput,
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorFieldView,
	InspectorSectionDescriptor,
	InspectorSectionView,
	LiveColorInput,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeColorControl, makeNumberField } from '../controls';
import { makeCheckboxField, makeRangeField, makeSelectField } from './controls-extra';

export interface InspectorFieldsOptions<T, P> {
	/** The section descriptors to draw (from `ooxml-ui/pptx`). */
	sections: readonly InspectorSectionDescriptor<T, P>[];
	/** Receives the patch a control change produces. */
	onPatch(patch: P): void;
	/**
	 * Receives a colour field's live patch while its picker is dragged: apply it
	 * WITHOUT an undo step. The drag's one `onPatch` follows when it settles.
	 */
	onPreview?(patch: P): void;
	/** A colour pick was committed (for the "Recent colours" list). */
	onColorCommit?(hex: string): void;
	/** Class of each section's box (default: the stacked `pptxv-inspector-fields-section`). */
	sectionClassName?: string;
}

export interface InspectorFieldsHandle<T> {
	el: HTMLElement;
	/** Repaint for the current value; rebuilds a section only when its shape changed. */
	update(target: T, disabled?: boolean): void;
}

interface FieldHandle {
	el: HTMLElement;
	setValue(value: InspectorFieldValue): void;
	setDisabled(disabled: boolean): void;
}

interface SectionSlot {
	el: HTMLElement;
	signature: string;
	toggle?: { setValue(on: boolean): void; setDisabled(disabled: boolean): void };
	fields: Map<string, FieldHandle>;
}

const signatureOf = (view: InspectorSectionView): string =>
	`${view.on}|${view.fields.map((field) => field.id).join(',')}`;

/** Put `testId` on the control inside a field row built by the shared factories. */
function tagControl(el: HTMLElement, testId: string | undefined): void {
	if (!testId) {
		return;
	}
	const control = el.querySelector('input, select, pptx-ui-select, pptx-ui-checkbox') ?? el;
	control.setAttribute('data-testid', testId);
}

/**
 * The vanilla binding's one generic field renderer: builds the DOM for any list
 * of shared inspector section descriptors (a toggle per section, then its
 * fields) out of the panel's own field factories, so every row carries the
 * panel's styling. Values, bounds, labels and patches all come from the
 * descriptors; `update` repaints values in place and rebuilds a section only
 * when its toggle or visible fields change, so a focused field is not replaced
 * while it is being edited.
 */
export function createInspectorFields<T, P>(
	doc: Document,
	t: Translator,
	options: InspectorFieldsOptions<T, P>,
): InspectorFieldsHandle<T> {
	const { sections } = options;
	const el = createEl(doc, 'div', 'pptxv-inspector-fields');
	const slots = new Map<string, SectionSlot>();
	let current: T;

	const emit = (patch: P | undefined): void => {
		if (patch !== undefined) {
			options.onPatch(patch);
		}
	};
	const change = (sectionId: string, fieldId: string, raw: InspectorFieldValue): void =>
		emit(inspectorFieldPatch(sections, sectionId, fieldId, current, raw));

	// One live controller per colour field (shared `createLiveColorPatchInput`):
	// a drag previews through `onPreview` and settles as one `onPatch`. Kept
	// outside the section slots, so a rebuilt section does not drop a drag.
	const colorInputs = new Map<string, LiveColorInput>();
	const colorInput = (sectionId: string, fieldId: string): LiveColorInput => {
		const key = `${sectionId}.${fieldId}`;
		let control = colorInputs.get(key);
		if (!control) {
			control = createLiveColorPatchInput<T, P & object>({
				read: () => {
					const field = resolveInspectorSections(sections, current)
						.find((view) => view.id === sectionId)
						?.fields.find((candidate) => candidate.id === fieldId);
					return field && { key, color: String(field.value), state: current };
				},
				patch: (target, hex) =>
					inspectorFieldPatch(sections, sectionId, fieldId, target, hex) as
						| (P & object)
						| undefined,
				preview: (patch) => options.onPreview?.(patch),
				commit: (patch) => options.onPatch(patch),
				remember: (hex) => options.onColorCommit?.(hex),
			});
			colorInputs.set(key, control);
		}
		return control;
	};

	const buildField = (sectionId: string, field: InspectorFieldView): FieldHandle => {
		const label = t(field.labelKey);
		let handle: FieldHandle;
		if (field.control === 'checkbox') {
			const box = makeCheckboxField(doc, {
				label,
				onChange: (checked) => change(sectionId, field.id, checked),
			});
			handle = { ...box, setValue: (v) => box.setValue(v === true) };
		} else if (field.control === 'select') {
			const select = makeSelectField<string>(doc, {
				label,
				options: (field.options ?? []).map((o) => ({ value: o.value, label: t(o.labelKey) })),
				onChange: (value) => change(sectionId, field.id, value),
			});
			handle = { ...select, setValue: (v) => select.setValue(String(v)) };
		} else if (field.control === 'color') {
			const row = createEl(doc, 'div', 'pptxv-inspector-row');
			const caption = createEl(doc, 'span', 'pptxv-inspector-row-label');
			caption.textContent = label;
			const live = colorInput(sectionId, field.id);
			const color = makeColorControl(
				doc,
				{
					label,
					onInput: (hex) => live.input(hex),
					onCommit: (hex) => live.change(hex),
					onBlur: (hex) => live.blur(hex),
				},
				String(field.value),
			);
			row.append(caption, color.el);
			handle = {
				el: row,
				setValue: (v) => color.setValue(String(v)),
				setDisabled: color.setDisabled,
			};
		} else if (field.control === 'range') {
			const range = makeRangeField(doc, {
				label,
				min: field.min ?? 0,
				max: field.max ?? 100,
				...(field.step !== undefined ? { step: field.step } : {}),
				formatValue: (v) => `${Math.round(v)}${field.unit ?? ''}`,
				onCommit: (v) => change(sectionId, field.id, v),
			});
			handle = { ...range, setValue: (v) => range.setValue(Number(v)) };
		} else {
			const number = makeNumberField(doc, {
				label,
				...(field.min !== undefined ? { min: field.min } : {}),
				...(field.max !== undefined ? { max: field.max } : {}),
				...(field.step !== undefined ? { step: field.step } : {}),
				onCommit: (v) => change(sectionId, field.id, v),
			});
			handle = { ...number, setValue: (v) => number.setValue(Number(v)) };
		}
		handle.el.classList.toggle('pptxv-inspector-fields-wide', field.wide);
		tagControl(handle.el, field.testId);
		return handle;
	};

	const buildSection = (view: InspectorSectionView): SectionSlot => {
		const box = createEl(doc, 'div', options.sectionClassName ?? 'pptxv-inspector-fields-section');
		box.dataset['section'] = view.id;
		const slot: SectionSlot = { el: box, signature: signatureOf(view), fields: new Map() };
		if (view.toggleable) {
			const toggle = makeCheckboxField(doc, {
				label: t(view.labelKey ?? ''),
				onChange: (on) => emit(inspectorTogglePatch(sections, view.id, current, on)),
			});
			toggle.el.classList.add('pptxv-inspector-fields-toggle');
			tagControl(toggle.el, view.toggleTestId);
			box.appendChild(toggle.el);
			slot.toggle = toggle;
		} else if (view.labelKey) {
			const title = createEl(doc, 'div', 'pptxv-inspector-fields-title');
			title.textContent = t(view.labelKey);
			box.appendChild(title);
		}
		for (const field of view.fields) {
			const handle = buildField(view.id, field);
			slot.fields.set(field.id, handle);
			box.appendChild(handle.el);
		}
		return slot;
	};

	return {
		el,
		update(target, disabled = false) {
			current = target;
			const views = resolveInspectorSections(sections, target);
			views.forEach((view, index) => {
				let slot = slots.get(view.id);
				if (!slot || slot.signature !== signatureOf(view)) {
					const next = buildSection(view);
					if (slot) {
						slot.el.replaceWith(next.el);
					} else {
						el.insertBefore(next.el, el.children[index] ?? null);
					}
					slots.set(view.id, next);
					slot = next;
				}
				slot.toggle?.setValue(view.on);
				slot.toggle?.setDisabled(disabled);
				for (const field of view.fields) {
					const handle = slot.fields.get(field.id);
					handle?.setValue(field.value);
					handle?.setDisabled(disabled);
				}
			});
		},
	};
}
