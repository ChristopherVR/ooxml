import {
	createInspectorColorFields,
	inspectorFieldPatch,
	inspectorOptionLabel,
	inspectorTogglePatch,
	resolveInspectorSections,
} from 'ooxml-ui/pptx';
import type {
	InspectorFieldValue,
	InspectorFieldView,
	InspectorSectionDescriptor,
	InspectorSectionView,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeColorControl, makeNumberField } from '../controls';
import {
	makeCheckboxField,
	makeRangeField,
	makeSelectField,
	makeTextField,
} from './controls-extra';

export interface InspectorFieldsOptions<T, P> {
	/** The section descriptors to draw (from `ooxml-ui/pptx`). */
	sections:
		| readonly InspectorSectionDescriptor<T, P>[]
		| ((target: T) => readonly InspectorSectionDescriptor<T, P>[]);
	/** Receives the patch a control change produces. */
	onPatch(patch: P): void;
	/**
	 * Receives a colour field's live patch while its picker is dragged, for the
	 * element `target` names (the current one, or one the drag has left): apply it
	 * WITHOUT an undo step. The drag's one `onPatch` follows when it settles.
	 */
	onPreview?(patch: P, target: string): void;
	/**
	 * What the descriptors read on element `key` now: a colour drag the selection has
	 * left is put back on it (none: such a drag is not put back). Not where `T` is
	 * inferred from: a lookup can always miss, a target need not.
	 */
	lookup?(key: string): NoInfer<T> | undefined;
	/** A colour pick was committed (for the "Recent colours" list). */
	onColorCommit?(hex: string): void;
	/** Class of each section's box (default: the stacked `pptxv-inspector-fields-section`). */
	sectionClassName?: string;
}

export interface InspectorFieldsHandle<T> {
	el: HTMLElement;
	/**
	 * Repaint for the current value; rebuilds a section only when its shape changed.
	 * `key` names the element `target` belongs to, so a colour field tells a new
	 * selection from a re-render of its own drag.
	 */
	update(target: T, disabled?: boolean, key?: string): void;
}

interface FieldHandle {
	el: HTMLElement;
	setValue(value: InspectorFieldValue): void;
	setDisabled(disabled: boolean): void;
	/** A gallery's preview picture, repainted in place. */
	setPreview?(src: string | undefined): void;
}

interface SectionSlot {
	el: HTMLElement;
	signature: string;
	toggle?: { setValue(on: boolean): void; setDisabled(disabled: boolean): void };
	fields: Map<string, FieldHandle>;
}

/** A section's shape: its toggle, its visible fields and their choices (resolved per target). */
const signatureOf = (view: InspectorSectionView): string =>
	`${view.on}|${view.label ?? view.labelKey ?? ''}|${view.fields
		.map((field) =>
			field.options
				? `${field.id}:${field.options.map((o) => `${o.value}=${o.label ?? ''}`).join(';')}`
				: field.id,
		)
		.join(',')}`;

/** Put `testId` on the control inside a field row built by the shared factories. */
function tagControl(el: HTMLElement, testId: string | undefined): void {
	if (!testId) {
		return;
	}
	const control =
		el.querySelector('input, textarea, select, pptx-ui-select, pptx-ui-checkbox') ?? el;
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
	const sectionsFor = (target: T) =>
		typeof options.sections === 'function' ? options.sections(target) : options.sections;
	const el = createEl(doc, 'div', 'pptxv-inspector-fields');
	const slots = new Map<string, SectionSlot>();
	let current: T;
	let currentKey = '';

	const emit = (patch: P | undefined): void => {
		if (patch !== undefined) {
			options.onPatch(patch);
		}
	};
	const change = (sectionId: string, fieldId: string, raw: InspectorFieldValue): void =>
		emit(inspectorFieldPatch(sectionsFor(current), sectionId, fieldId, current, raw));

	// One live controller per colour field (shared `createInspectorColorFields`):
	// a drag previews through `onPreview` and settles as one `onPatch`. Kept
	// outside the section slots, so a rebuilt section does not drop a drag.
	const colorInput = createInspectorColorFields<T, P>({
		sections: () => sectionsFor(current),
		target: () => current,
		targetKey: () => currentKey,
		lookup: (key) => options.lookup?.(key),
		preview: (patch, target) => options.onPreview?.(patch, target),
		commit: (patch) => options.onPatch(patch),
		remember: (hex) => options.onColorCommit?.(hex),
	});

	const buildField = (sectionId: string, field: InspectorFieldView): FieldHandle => {
		const label = t(field.labelKey);
		let handle: FieldHandle;
		if (field.control === 'gallery') {
			handle = buildGallery(sectionId, field, label);
		} else if (field.control === 'checkbox') {
			const box = makeCheckboxField(doc, {
				label,
				onChange: (checked) => change(sectionId, field.id, checked),
			});
			handle = { ...box, setValue: (v) => box.setValue(v === true) };
		} else if (field.control === 'select') {
			const select = makeSelectField<string>(doc, {
				label,
				options: (field.options ?? []).map((o) => ({
					value: o.value,
					label: inspectorOptionLabel(o, t),
				})),
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
		} else if (field.control === 'text' || field.control === 'textarea') {
			const text = makeTextField(doc, {
				label,
				multiline: field.control === 'textarea',
				...(field.placeholderKey ? { placeholder: t(field.placeholderKey) } : {}),
				onInput: (value) => change(sectionId, field.id, value),
			});
			handle = { ...text, setValue: (v) => text.setValue(String(v)) };
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
		if (field.control !== 'gallery') {
			tagControl(handle.el, field.testId);
		}
		return handle;
	};

	/** A grid of choices, each previewing the field's picture through its CSS filter. */
	const buildGallery = (
		sectionId: string,
		field: InspectorFieldView,
		label: string,
	): FieldHandle => {
		const box = createEl(doc, 'div', 'pptxv-inspector-gallery');
		box.setAttribute('role', 'group');
		box.setAttribute('aria-label', label);
		const caption = createEl(doc, 'span', 'pptxv-inspector-row-label');
		caption.textContent = label;
		const grid = createEl(doc, 'div', 'pptxv-inspector-gallery-grid');
		if (field.testId) {
			grid.setAttribute('data-testid', field.testId);
		}
		const items = (field.options ?? []).map((option) => {
			const text = inspectorOptionLabel(option, t);
			const button = createEl(doc, 'button', 'pptxv-inspector-gallery-item');
			button.type = 'button';
			button.title = text;
			button.setAttribute('aria-label', text);
			button.dataset['value'] = option.value;
			const preview = createEl(doc, 'span', 'pptxv-inspector-gallery-preview');
			if (option.previewFilter) {
				preview.style.filter = option.previewFilter;
			}
			const name = createEl(doc, 'span', 'pptxv-inspector-gallery-caption');
			name.textContent = text;
			button.append(preview, name);
			button.addEventListener('click', () => change(sectionId, field.id, option.value));
			grid.appendChild(button);
			return { button, preview, value: option.value };
		});
		box.append(caption, grid);
		const setPreview = (src: string | undefined): void => {
			for (const { preview } of items) {
				preview.style.backgroundImage = src ? `url("${src.replace(/"/g, '%22')}")` : '';
			}
		};
		setPreview(field.previewImage);
		return {
			el: box,
			setValue: (v) => {
				for (const { button, value } of items) {
					const active = value === String(v);
					button.classList.toggle('is-active', active);
					button.setAttribute('aria-pressed', String(active));
				}
			},
			setDisabled: (disabled) => {
				for (const { button } of items) {
					button.disabled = disabled;
				}
			},
			setPreview,
		};
	};

	const buildSection = (view: InspectorSectionView): SectionSlot => {
		const box = createEl(doc, 'div', options.sectionClassName ?? 'pptxv-inspector-fields-section');
		box.dataset['section'] = view.id;
		const slot: SectionSlot = { el: box, signature: signatureOf(view), fields: new Map() };
		if (view.toggleable) {
			const toggle = makeCheckboxField(doc, {
				label: view.label ?? t(view.labelKey ?? ''),
				onChange: (on) => emit(inspectorTogglePatch(sectionsFor(current), view.id, current, on)),
			});
			toggle.el.classList.add('pptxv-inspector-fields-toggle');
			tagControl(toggle.el, view.toggleTestId);
			box.appendChild(toggle.el);
			slot.toggle = toggle;
		} else if (view.label || view.labelKey) {
			const title = createEl(doc, 'div', 'pptxv-inspector-fields-title');
			title.textContent = view.label ?? t(view.labelKey ?? '');
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
		update(target, disabled = false, key = '') {
			current = target;
			currentKey = key;
			const views = resolveInspectorSections(sectionsFor(target), target);
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
					handle?.setPreview?.(field.previewImage);
				}
			});
		},
	};
}
