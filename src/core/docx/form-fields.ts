// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Legacy form fields (Word 97 text boxes, check boxes and drop-downs): reading `w:ffData` into the
// model, writing an edited field back over its source element, and the DOM-free edits a click or
// a pick makes. A text field's value is its result text; a check box or drop-down has no result
// text, so its value lives in `w:ffData` and is drawn from there.
import { isStFFTextType, isStInfoTextType } from './generated/wml-simple-types';
import type { FormField, FormFieldInfoText } from './form-field-model';
import type { TextRun } from './model';
import { onOffElement, parseInteger, parseUnsignedInteger } from './simple-types';
import {
	buildXml,
	first,
	getW,
	isElement,
	makeW,
	named,
	parseXml,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';

const KINDS = { textInput: 'text', checkBox: 'checkbox', ddList: 'dropDown' } as const;
type KindElement = keyof typeof KINDS;
const elementOfKind = (kind: FormField['kind']): KindElement =>
	(Object.keys(KINDS) as KindElement[]).find((name) => KINDS[name] === kind)!;

const elementChildren = (parent: XmlElement): XmlElement[] =>
	Array.from(parent.childNodes).filter(isElement);

function infoText(element: XmlElement | undefined): FormFieldInfoText | undefined {
	if (!element) return undefined;
	const type = getW(element, 'type') ?? 'text';
	return isStInfoTextType(type) ? { type, value: getW(element, 'val') ?? '' } : undefined;
}

/** The form field `w:ffData` describes, or undefined when it holds no text, check box or list. */
export function parseFormField(ffData: XmlElement): FormField | undefined {
	const typed = elementChildren(ffData).find((child) =>
		Object.hasOwn(KINDS, child.localName ?? ''),
	);
	if (!typed || typed.namespaceURI !== WORD_NS) return undefined;
	const field: FormField = { kind: KINDS[typed.localName as KindElement] };
	const name = getW(first(ffData, 'name'), 'val');
	if (name) field.name = name;
	const enabled = onOffElement(first(ffData, 'enabled'));
	if (enabled !== undefined) field.enabled = enabled;
	const calcOnExit = onOffElement(first(ffData, 'calcOnExit'));
	if (calcOnExit !== undefined) field.calcOnExit = calcOnExit;
	const helpText = infoText(first(ffData, 'helpText'));
	if (helpText) field.helpText = helpText;
	const statusText = infoText(first(ffData, 'statusText'));
	if (statusText) field.statusText = statusText;
	if (field.kind === 'text') {
		const type = getW(first(typed, 'type'), 'val');
		const fallback = getW(first(typed, 'default'), 'val');
		const maxLength = parseUnsignedInteger(getW(first(typed, 'maxLength'), 'val'));
		const format = getW(first(typed, 'format'), 'val');
		field.textInput = {
			...(isStFFTextType(type) && { type }),
			...(fallback !== undefined && { default: fallback }),
			...(maxLength !== undefined && { maxLength }),
			...(format !== undefined && { format }),
		};
	} else if (field.kind === 'checkbox') {
		const checked = onOffElement(first(typed, 'checked'));
		const sizeAuto = onOffElement(first(typed, 'sizeAuto'));
		const size = parseUnsignedInteger(getW(first(typed, 'size'), 'val'));
		field.checkBox = {
			default: onOffElement(first(typed, 'default')) ?? false,
			...(checked !== undefined && { checked }),
			...(sizeAuto !== undefined && { sizeAuto }),
			...(size !== undefined && { size }),
		};
	} else {
		const result = parseInteger(getW(first(typed, 'result'), 'val'));
		const fallback = parseInteger(getW(first(typed, 'default'), 'val'));
		field.dropDown = {
			entries: elementChildren(typed)
				.filter((child) => named(child, 'listEntry'))
				.map((entry) => getW(entry, 'val') ?? ''),
			...(result !== undefined && { result }),
			...(fallback !== undefined && { default: fallback }),
		};
	}
	field.sourceXml = buildXml(ffData);
	return field;
}

/** Whether a `w:fldChar` child is form field data this model reads and writes back. */
export const isModeledFormFieldData = (element: XmlElement): boolean =>
	named(element, 'ffData') && parseFormField(element) !== undefined;

const withoutSource = ({ sourceXml: _source, ...field }: FormField) => field;
const sameField = (a: FormField, b: FormField) =>
	JSON.stringify(withoutSource(a)) === JSON.stringify(withoutSource(b));

function valued(doc: XmlDocument, local: string, value: string): XmlElement {
	const element = makeW(doc, local);
	element.setAttributeNS(WORD_NS, 'w:val', value);
	return element;
}

/** `<w:x/>` for on, `<w:x w:val="0"/>` for off. */
const onOff = (doc: XmlDocument, local: string, on: boolean): XmlElement =>
	on ? makeW(doc, local) : valued(doc, local, '0');

function kindElement(doc: XmlDocument, field: FormField): XmlElement {
	const element = makeW(doc, elementOfKind(field.kind));
	const add = (child: XmlElement | undefined) => child && element.appendChild(child);
	if (field.kind === 'text') {
		const input = field.textInput ?? {};
		if (input.type) add(valued(doc, 'type', input.type));
		if (input.default !== undefined) add(valued(doc, 'default', input.default));
		if (input.maxLength !== undefined) add(valued(doc, 'maxLength', String(input.maxLength)));
		if (input.format !== undefined) add(valued(doc, 'format', input.format));
	} else if (field.kind === 'checkbox') {
		const box = field.checkBox ?? { default: false };
		add(
			box.size !== undefined
				? valued(doc, 'size', String(box.size))
				: onOff(doc, 'sizeAuto', box.sizeAuto ?? true),
		);
		add(valued(doc, 'default', box.default ? '1' : '0'));
		if (box.checked !== undefined) add(onOff(doc, 'checked', box.checked));
	} else {
		const list = field.dropDown ?? { entries: [] };
		if (list.result !== undefined) add(valued(doc, 'result', String(list.result)));
		if (list.default !== undefined) add(valued(doc, 'default', String(list.default)));
		for (const entry of list.entries) add(valued(doc, 'listEntry', entry));
	}
	return element;
}

/** Replaces the `local` child of `ffData` in place, appends it before the kind element, or removes it. */
function setChild(ffData: XmlElement, local: string, next: XmlElement | undefined): void {
	const current = first(ffData, local);
	if (current && next) ffData.replaceChild(next, current);
	else if (current) ffData.removeChild(current);
	else if (next) {
		const kind = elementChildren(ffData).find((child) =>
			Object.hasOwn(KINDS, child.localName ?? ''),
		);
		ffData.insertBefore(next, kind ?? null);
	}
}

/**
 * The `w:ffData` for `field`: `source` (the field's own element from the file, when it is at
 * hand) or its verbatim XML, with what the model changed written over it.
 */
export function writeFormFieldData(
	doc: XmlDocument,
	field: FormField,
	source?: XmlElement,
): XmlElement {
	const ffData =
		source && buildXml(source) === field.sourceXml
			? source
			: field.sourceXml
				? (doc.importNode(parseXml(field.sourceXml).documentElement, true) as XmlElement)
				: makeW(doc, 'ffData');
	const current = parseFormField(ffData);
	if (current && sameField(current, field)) return ffData;
	if (current?.name !== field.name) setChild(ffData, 'name', valued(doc, 'name', field.name ?? ''));
	for (const key of ['enabled', 'calcOnExit'] as const)
		if (current?.[key] !== field[key])
			setChild(ffData, key, field[key] === undefined ? undefined : onOff(doc, key, field[key]));
	for (const key of ['helpText', 'statusText'] as const) {
		if (JSON.stringify(current?.[key]) === JSON.stringify(field[key])) continue;
		const text = field[key];
		const element = text && makeW(doc, key);
		if (element && text) {
			element.setAttributeNS(WORD_NS, 'w:type', text.type);
			element.setAttributeNS(WORD_NS, 'w:val', text.value);
		}
		setChild(ffData, key, element);
	}
	const kind = elementChildren(ffData).find((child) => Object.hasOwn(KINDS, child.localName ?? ''));
	if (kind) ffData.replaceChild(kindElement(doc, field), kind);
	else ffData.appendChild(kindElement(doc, field));
	return ffData;
}

/** Whether a check box form field shows as checked: its own state, else its default. */
export const formCheckboxChecked = (field: FormField): boolean =>
	field.checkBox?.checked ?? field.checkBox?.default ?? false;

/** The entry a drop-down form field shows: the selected one, else the default, else the first. */
export function formDropDownIndex(field: FormField): number | undefined {
	const entries = field.dropDown?.entries ?? [];
	for (const index of [field.dropDown?.result, field.dropDown?.default, 0])
		if (index !== undefined && index >= 0 && index < entries.length) return index;
	return undefined;
}

/**
 * What a check box or drop-down form field's begin marker shows, since neither has result text:
 * a ballot box, or the shown entry. Undefined for other runs, including text fields.
 */
export function formFieldDisplayText(run: TextRun): string | undefined {
	const field = run.fieldChar === 'begin' ? run.formField : undefined;
	if (field?.kind === 'checkbox') return formCheckboxChecked(field) ? '☒' : '☐';
	if (field?.kind === 'dropDown') {
		const index = formDropDownIndex(field);
		return index === undefined ? '' : field.dropDown!.entries[index];
	}
	return undefined;
}

/** The check box with `checked` as its state. */
export function setFormCheckbox(field: FormField, checked: boolean): FormField {
	return { ...field, checkBox: { ...(field.checkBox ?? { default: false }), checked } };
}

/** The drop-down with entry `index` selected, or undefined when there is no such entry. */
export function selectFormDropDownEntry(field: FormField, index: number): FormField | undefined {
	const list = field.dropDown;
	if (!list || !Number.isInteger(index) || index < 0 || index >= list.entries.length)
		return undefined;
	return { ...field, dropDown: { ...list, result: index } };
}

/**
 * The form field a run belongs to: the begin marker's own field, or the field whose begin marker
 * opens the innermost complex field around run `index`.
 */
export function formFieldOfRun(runs: readonly TextRun[], index: number): FormField | undefined {
	let depth = 0;
	for (let at = index; at >= 0; at--) {
		const run = runs[at];
		if (run?.fieldChar === 'end' && at !== index) depth++;
		if (run?.fieldChar !== 'begin') continue;
		if (depth === 0) return run.formField;
		depth--;
	}
	return undefined;
}
