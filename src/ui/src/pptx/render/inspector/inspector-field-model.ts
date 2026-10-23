/**
 * The inspector's field descriptor model: one framework-neutral description of
 * a Format pane section (its optional on/off toggle and its fields), written
 * once here and drawn by ONE generic field renderer per binding.
 *
 * A descriptor is generic over the value it edits (`T`, e.g. a `ShapeStyle` or
 * a `PptxElement`) and the patch a change produces (`P`, e.g. a
 * `Partial<ShapeStyle>`), so the same model serves every section. A binding
 * never reads or writes the model itself: it asks {@link resolveInspectorSections}
 * for the current views (labels, bounds, values) and turns a control event into
 * a patch with {@link inspectorFieldPatch} / {@link inspectorTogglePatch}.
 *
 * @module render/inspector/inspector-field-model
 */

/**
 * What a field edits, which decides its control and how input is coerced:
 * - `number`, `length`: a number box, clamped to `min`/`max`.
 * - `angle`: a number box whose value wraps into `[0, 360)`.
 * - `percent`: a 0-100 slider, clamped.
 * - `color`: a colour picker; the value is a `#rrggbb` string.
 * - `select`: a drop-down of `options`.
 * - `toggle`: a checkbox.
 */
export type InspectorFieldKind =
	| 'number'
	| 'length'
	| 'angle'
	| 'percent'
	| 'color'
	| 'select'
	| 'toggle';

/** The control a binding draws for a field kind. */
export type InspectorFieldControl = 'number' | 'range' | 'color' | 'select' | 'checkbox';

/** A value a field reads or a control reports. */
export type InspectorFieldValue = number | string | boolean;

/** One choice of a `select` field. */
export interface InspectorFieldOption {
	value: string;
	/** Translation key of the choice's label. */
	labelKey: string;
}

/** One editable field of a section. */
export interface InspectorFieldDescriptor<T, P> {
	/** Stable id within its section; what a change names. */
	id: string;
	kind: InspectorFieldKind;
	/** Translation key of the field's caption (also its accessible name). */
	labelKey: string;
	min?: number;
	max?: number;
	step?: number;
	/** Unit suffix a binding may show ("pt", "%", "°"). */
	unit?: string;
	/** Spans the full width of a two-column field grid. */
	wide?: boolean;
	/** `data-testid` of the control. */
	testId?: string;
	options?: readonly InspectorFieldOption[];
	/** Hide the field for this value (the field is shown when absent). */
	visibleWhen?: (target: T) => boolean;
	/** The field's current value. */
	read: (target: T) => InspectorFieldValue;
	/** The patch that sets the field to an already coerced `value`. */
	write: (target: T, value: InspectorFieldValue) => P;
}

/** A section's on/off switch (Shadow, Glow, ...); its fields show only when on. */
export interface InspectorSectionToggle<T, P> {
	isOn: (target: T) => boolean;
	enable: (target: T) => P;
	disable: (target: T) => P;
	/** `data-testid` of the checkbox. */
	testId?: string;
}

/** A titled group of fields, optionally behind a toggle. */
export interface InspectorSectionDescriptor<T, P> {
	id: string;
	/** Translation key of the section title (the toggle's caption when toggleable). */
	labelKey: string;
	toggle?: InspectorSectionToggle<T, P>;
	fields: readonly InspectorFieldDescriptor<T, P>[];
}

/** What a binding draws for one field. */
export interface InspectorFieldView {
	id: string;
	kind: InspectorFieldKind;
	control: InspectorFieldControl;
	labelKey: string;
	value: InspectorFieldValue;
	min?: number;
	max?: number;
	step?: number;
	unit?: string;
	wide: boolean;
	testId?: string;
	options?: readonly InspectorFieldOption[];
}

/** What a binding draws for one section. */
export interface InspectorSectionView {
	id: string;
	labelKey: string;
	/** The section has an on/off checkbox. */
	toggleable: boolean;
	/** The toggle is on (always `true` for a section without one). */
	on: boolean;
	toggleTestId?: string;
	/** The visible fields; empty while a toggle is off. */
	fields: InspectorFieldView[];
}

const CONTROL_OF: Record<InspectorFieldKind, InspectorFieldControl> = {
	number: 'number',
	length: 'number',
	angle: 'number',
	percent: 'range',
	color: 'color',
	select: 'select',
	toggle: 'checkbox',
};

/** Display rounding for numeric values (two decimals, as PowerPoint's panes show). */
function displayNumber(value: InspectorFieldValue): InspectorFieldValue {
	return typeof value === 'number' && Number.isFinite(value)
		? Math.round(value * 100) / 100
		: value;
}

function fieldView<T, P>(field: InspectorFieldDescriptor<T, P>, target: T): InspectorFieldView {
	const view: InspectorFieldView = {
		id: field.id,
		kind: field.kind,
		control: CONTROL_OF[field.kind],
		labelKey: field.labelKey,
		value: displayNumber(field.read(target)),
		wide: field.wide === true,
	};
	if (field.kind === 'percent') {
		view.min = field.min ?? 0;
		view.max = field.max ?? 100;
	} else {
		if (field.min !== undefined) view.min = field.min;
		if (field.max !== undefined) view.max = field.max;
	}
	if (field.step !== undefined) view.step = field.step;
	if (field.unit !== undefined) view.unit = field.unit;
	if (field.testId !== undefined) view.testId = field.testId;
	if (field.options !== undefined) view.options = field.options;
	return view;
}

/** Resolve the sections' current views for `target`. */
export function resolveInspectorSections<T, P>(
	sections: readonly InspectorSectionDescriptor<T, P>[],
	target: T,
): InspectorSectionView[] {
	return sections.map((section) => {
		const on = section.toggle ? section.toggle.isOn(target) : true;
		const view: InspectorSectionView = {
			id: section.id,
			labelKey: section.labelKey,
			toggleable: section.toggle !== undefined,
			on,
			fields: on
				? section.fields
						.filter((field) => field.visibleWhen?.(target) ?? true)
						.map((field) => fieldView(field, target))
				: [],
		};
		if (section.toggle?.testId !== undefined) view.toggleTestId = section.toggle.testId;
		return view;
	});
}

/**
 * Coerce a raw control value (an input's string, a checkbox's boolean) for
 * `field`: numbers are parsed and clamped (angles wrap), toggles become
 * booleans, colours and choices stay strings. `undefined` when the input is
 * not a usable value (an empty or non-numeric box), so nothing is written.
 */
export function coerceInspectorFieldValue<T, P>(
	field: InspectorFieldDescriptor<T, P>,
	raw: InspectorFieldValue,
): InspectorFieldValue | undefined {
	switch (field.kind) {
		case 'toggle':
			return typeof raw === 'boolean' ? raw : raw === 'true';
		case 'color':
		case 'select': {
			const text = String(raw).trim();
			return text.length > 0 ? text : undefined;
		}
		default: {
			if (typeof raw === 'string' && raw.trim() === '') {
				return undefined;
			}
			const n = typeof raw === 'number' ? raw : Number(raw);
			if (!Number.isFinite(n)) {
				return undefined;
			}
			if (field.kind === 'angle') {
				return ((n % 360) + 360) % 360;
			}
			const lo = field.min ?? (field.kind === 'percent' ? 0 : -Infinity);
			const hi = field.max ?? (field.kind === 'percent' ? 100 : Infinity);
			return Math.max(lo, Math.min(hi, n));
		}
	}
}

function findSection<T, P>(
	sections: readonly InspectorSectionDescriptor<T, P>[],
	sectionId: string,
): InspectorSectionDescriptor<T, P> | undefined {
	return sections.find((section) => section.id === sectionId);
}

/**
 * The patch for a control change: `raw` is the control's value. `undefined`
 * when the section or field is unknown or the input is unusable.
 */
export function inspectorFieldPatch<T, P>(
	sections: readonly InspectorSectionDescriptor<T, P>[],
	sectionId: string,
	fieldId: string,
	target: T,
	raw: InspectorFieldValue,
): P | undefined {
	const field = findSection(sections, sectionId)?.fields.find((f) => f.id === fieldId);
	if (!field) {
		return undefined;
	}
	const value = coerceInspectorFieldValue(field, raw);
	return value === undefined ? undefined : field.write(target, value);
}

/** The patch for switching a section's toggle; `undefined` without one. */
export function inspectorTogglePatch<T, P>(
	sections: readonly InspectorSectionDescriptor<T, P>[],
	sectionId: string,
	target: T,
	on: boolean,
): P | undefined {
	const toggle = findSection(sections, sectionId)?.toggle;
	if (!toggle) {
		return undefined;
	}
	return on ? toggle.enable(target) : toggle.disable(target);
}
