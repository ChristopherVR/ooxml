// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes content control (`w:sdt`) properties: patches only what the model changed on an existing
// control, and builds a new control (from its verbatim `w:sdtPr` when it has one).
import type {
	ContentControl,
	ContentControlGlyph,
	ContentControlListItem,
} from './content-control-model';
import { parseContentControl, propertiesKind, sdtContent, typeElements } from './content-control';
import {
	NAMED_TYPES,
	validId,
	validLock,
	verbatimEndProperties,
	verbatimProperties,
} from './content-control-values';
import { declareIgnorableExtension, WORD_2010_NS } from './ligatures';
import {
	children,
	first,
	getW,
	makeW,
	named,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';

/** `CT_SdtPr` child order; the type element (and everything unknown) follows these. */
const PROPERTY_ORDER = [
	'rPr',
	'alias',
	'tag',
	'id',
	'lock',
	'placeholder',
	'temporary',
	'showingPlcHdr',
	'dataBinding',
	'label',
	'tabIndex',
];
const rank = (element: XmlElement): number => {
	const index = element.namespaceURI === WORD_NS ? PROPERTY_ORDER.indexOf(element.localName) : -1;
	return index < 0 ? PROPERTY_ORDER.length : index;
};

function place(sdtPr: XmlElement, element: XmlElement): void {
	const siblings = Array.from(sdtPr.childNodes).filter((n): n is XmlElement => n.nodeType === 1);
	const after = siblings.find((sibling) => rank(sibling) > rank(element));
	sdtPr.insertBefore(element, after ?? null);
}

function setValued(doc: XmlDocument, sdtPr: XmlElement, local: string, value: string | undefined) {
	for (const existing of children(sdtPr, local)) sdtPr.removeChild(existing);
	if (value === undefined) return;
	const element = makeW(doc, local);
	element.setAttributeNS(WORD_NS, 'w:val', value);
	place(sdtPr, element);
}

const same = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);

function w14(doc: XmlDocument, local: string): XmlElement {
	const prefix = declareIgnorableExtension(doc, WORD_2010_NS, 'w14');
	return doc.createElementNS(WORD_2010_NS, `${prefix}:${local}`);
}
function w14Attribute(element: XmlElement, local: string, value: string): void {
	const prefix = element.lookupPrefix(WORD_2010_NS) ?? 'w14';
	element.setAttributeNS(WORD_2010_NS, `${prefix}:${local}`, value);
}

function checkboxElement(sdtPr: XmlElement): XmlElement | undefined {
	return Array.from(sdtPr.childNodes).find(
		(n): n is XmlElement =>
			n.nodeType === 1 &&
			(n as XmlElement).localName === 'checkbox' &&
			(n as XmlElement).namespaceURI === WORD_2010_NS,
	);
}

// A list, a check box or a date pasted from another page holds whatever that page wrote: only
// values Word reads are written.

/** The entries of a list control that can be written, each with a text display and value. */
function validItems(items: unknown): ContentControlListItem[] {
	if (!Array.isArray(items)) return [];
	return items.flatMap((item: unknown) => {
		if (!item || typeof item !== 'object') return [];
		const { display, value } = item as { display?: unknown; value?: unknown };
		const text = display ?? value;
		return text === undefined || text === null
			? []
			: [{ display: String(text), value: String(value ?? text) }];
	});
}

/** A check box glyph Word reads: a code point in four hexadecimal digits (`ST_ShortHexNumber`). */
function validGlyph(glyph: unknown): ContentControlGlyph | undefined {
	const { code, font } = (glyph ?? {}) as { code?: unknown; font?: unknown };
	if (typeof code !== 'string' || !/^[0-9A-Fa-f]{4}$/.test(code)) return undefined;
	return typeof font === 'string' ? { code, font } : { code };
}

/** A `w:fullDate`, an `xsd:dateTime`; anything else is not written. */
const validDate = (value: unknown): string | undefined =>
	typeof value === 'string' &&
	/^-?\d{4,}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})?$/.test(value)
		? value
		: undefined;

function writeCheckbox(doc: XmlDocument, sdtPr: XmlElement, control: ContentControl): void {
	let box = checkboxElement(sdtPr);
	if (!box) {
		box = w14(doc, 'checkbox');
		sdtPr.appendChild(box);
	}
	const replace = (local: string, build: (element: XmlElement) => void, order: number) => {
		for (const old of Array.from(box!.childNodes))
			if ((old as XmlElement).localName === local) box!.removeChild(old);
		const element = w14(doc, local);
		build(element);
		const siblings = Array.from(box!.childNodes);
		box!.insertBefore(element, siblings[order] ?? null);
	};
	replace('checked', (el) => w14Attribute(el, 'val', control.checked ? '1' : '0'), 0);
	const states = [validGlyph(control.checkedGlyph), validGlyph(control.uncheckedGlyph)];
	states.forEach((glyph, index) => {
		if (glyph)
			replace(
				index ? 'uncheckedState' : 'checkedState',
				(el) => {
					w14Attribute(el, 'val', glyph.code);
					if (glyph.font) w14Attribute(el, 'font', glyph.font);
				},
				index + 1,
			);
	});
}

function writeItems(doc: XmlDocument, sdtPr: XmlElement, control: ContentControl): void {
	const list = first(sdtPr, control.kind);
	if (!list) return;
	for (const item of children(list, 'listItem')) list.removeChild(item);
	for (const item of validItems(control.items)) {
		const element = makeW(doc, 'listItem');
		element.setAttributeNS(WORD_NS, 'w:displayText', item.display);
		element.setAttributeNS(WORD_NS, 'w:value', item.value);
		list.appendChild(element);
	}
}

function writeDate(doc: XmlDocument, sdtPr: XmlElement, control: ContentControl): void {
	const date = first(sdtPr, 'date');
	if (!date) return;
	const fullDate = validDate(control.fullDate);
	if (fullDate) date.setAttributeNS(WORD_NS, 'w:fullDate', fullDate);
	else date.removeAttributeNS(WORD_NS, 'fullDate');
	const set = (local: string, value: string | undefined, order: number) => {
		for (const old of children(date, local)) date.removeChild(old);
		if (value === undefined) return;
		const element = makeW(doc, local);
		element.setAttributeNS(WORD_NS, 'w:val', String(value));
		date.insertBefore(element, Array.from(date.childNodes)[order] ?? null);
	};
	set('dateFormat', control.dateFormat, 0);
	set('lid', control.dateLocale, 1);
}

/** Whether the writer writes a type element for `kind`: rich text has none, unknown kinds none. */
const writesType = (kind: string): boolean =>
	NAMED_TYPES.has(kind) ||
	['richText', 'text', 'checkbox', 'dropDownList', 'comboBox', 'date'].includes(kind);

/** Applies every modeled field of `control` that differs from what `sdt`'s own `w:sdtPr` says. */
export function patchContentControl(doc: XmlDocument, sdt: XmlElement, control: ContentControl) {
	const sdtPr = first(sdt, 'sdtPr');
	if (!sdtPr) return;
	const current = parseContentControl(sdt, control.level);
	// The model can hold another control than this `w:sdt` (a row or cell given another one): its
	// id and kind are written too, the old kind's element going, so the properties stay one control.
	if (control.id !== current.id && validId(control.id) !== undefined)
		setValued(doc, sdtPr, 'id', control.id);
	if (control.kind !== current.kind && writesType(control.kind)) {
		for (const old of typeElements(sdtPr)) sdtPr.removeChild(old);
		const { sourceXml: _source, ...model } = control;
		for (const type of typeElements(buildProperties(doc, model))) sdtPr.appendChild(type);
	}
	if (control.alias !== current.alias) setValued(doc, sdtPr, 'alias', control.alias);
	if (control.tag !== current.tag) setValued(doc, sdtPr, 'tag', control.tag);
	const lock = validLock(control.lock);
	if (lock !== current.lock) setValued(doc, sdtPr, 'lock', lock);
	if (Boolean(control.showingPlaceholder) !== Boolean(current.showingPlaceholder)) {
		for (const old of children(sdtPr, 'showingPlcHdr')) sdtPr.removeChild(old);
		if (control.showingPlaceholder) place(sdtPr, makeW(doc, 'showingPlcHdr'));
	}
	if (Boolean(control.temporary) !== Boolean(current.temporary)) {
		for (const old of children(sdtPr, 'temporary')) sdtPr.removeChild(old);
		if (control.temporary) place(sdtPr, makeW(doc, 'temporary'));
	}
	if (control.kind === 'checkbox') {
		if (
			Boolean(control.checked) !== Boolean(current.checked) ||
			!same(control.checkedGlyph, current.checkedGlyph) ||
			!same(control.uncheckedGlyph, current.uncheckedGlyph)
		)
			writeCheckbox(doc, sdtPr, control);
	}
	if (
		(control.kind === 'dropDownList' || control.kind === 'comboBox') &&
		!same(validItems(control.items), current.items)
	)
		writeItems(doc, sdtPr, control);
	if (
		control.kind === 'date' &&
		(control.fullDate !== current.fullDate ||
			control.dateFormat !== current.dateFormat ||
			control.dateLocale !== current.dateLocale)
	)
		writeDate(doc, sdtPr, control);
}

/**
 * Builds the `w:sdtPr` of a control that has no source element: its verbatim one when it has a
 * `w:sdtPr`, else one from the model.
 */
function buildProperties(doc: XmlDocument, control: ContentControl): XmlElement {
	const verbatim = verbatimProperties(doc, control.sourceXml);
	if (verbatim) return verbatim;
	const sdtPr = makeW(doc, 'sdtPr');
	if (control.alias !== undefined) setValued(doc, sdtPr, 'alias', control.alias);
	if (control.tag !== undefined) setValued(doc, sdtPr, 'tag', control.tag);
	setValued(doc, sdtPr, 'id', validId(control.id));
	setValued(doc, sdtPr, 'lock', validLock(control.lock));
	if (control.showingPlaceholder) place(sdtPr, makeW(doc, 'showingPlcHdr'));
	if (control.temporary) place(sdtPr, makeW(doc, 'temporary'));
	switch (control.kind) {
		case 'text': {
			const text = makeW(doc, 'text');
			if (control.multiLine) text.setAttributeNS(WORD_NS, 'w:multiLine', '1');
			sdtPr.appendChild(text);
			break;
		}
		case 'checkbox':
			writeCheckbox(doc, sdtPr, control);
			break;
		case 'dropDownList':
		case 'comboBox':
			sdtPr.appendChild(makeW(doc, control.kind));
			writeItems(doc, sdtPr, control);
			break;
		case 'date':
			sdtPr.appendChild(makeW(doc, 'date'));
			writeDate(doc, sdtPr, control);
			break;
		default:
			if (NAMED_TYPES.has(control.kind)) sdtPr.appendChild(makeW(doc, control.kind));
	}
	return sdtPr;
}

/** A new `w:sdt` for `control` with an empty `w:sdtContent`. */
export function createContentControl(
	doc: XmlDocument,
	control: ContentControl,
): { sdt: XmlElement; content: XmlElement } {
	const sdt = makeW(doc, 'sdt');
	// Verbatim properties of another kind are another control's: the model's kind wins.
	const { sourceXml, ...model } = control;
	const sdtPr = buildProperties(
		doc,
		sourceXml && propertiesKind(sourceXml) !== control.kind ? model : control,
	);
	// A copy keeps the properties it was copied with, under the new id it was given: ids are
	// unique in a document (ECMA-376 17.5.2.18). An id Word cannot read, such as one generated for
	// a control read without `w:id`, is not written, nor the one its properties were copied with.
	if (getW(first(sdtPr, 'id'), 'val') !== control.id)
		setValued(doc, sdtPr, 'id', validId(control.id));
	sdt.appendChild(sdtPr);
	const end = verbatimEndProperties(doc, control.endSourceXml);
	if (end) sdt.appendChild(end);
	const content = makeW(doc, 'sdtContent');
	sdt.appendChild(content);
	patchContentControl(doc, sdt, control);
	return { sdt, content };
}

/**
 * Replaces a `w:sdt` by its content: the control goes, what it held stays where it was, and so do
 * the range markers (bookmarks, comment, custom XML and move ranges) that Word writes beside the
 * content of a row or cell control (MS-OI29500 2.1.202, 2.1.204), before or after it as they were.
 */
export function unwrapContentControl(sdt: XmlElement): void {
	const parent = sdt.parentNode;
	if (!parent) return;
	for (const child of Array.from(sdt.childNodes)) {
		if (child.nodeType !== 1) continue;
		const element = child as XmlElement;
		if (named(element, 'sdtContent'))
			for (const inner of Array.from(element.childNodes)) parent.insertBefore(inner, sdt);
		else if (!named(element, 'sdtPr') && !named(element, 'sdtEndPr'))
			parent.insertBefore(element, sdt);
	}
	parent.removeChild(sdt);
}

/** Removes the children of a reused `w:sdt`'s content and returns it, ready for new runs. */
export function resetContentControl(
	doc: XmlDocument,
	sdt: XmlElement,
	control: ContentControl,
): XmlElement {
	patchContentControl(doc, sdt, control);
	let content = sdtContent(sdt);
	if (!content) {
		content = makeW(doc, 'sdtContent');
		sdt.appendChild(content);
	}
	for (const child of Array.from(content.childNodes)) content.removeChild(child);
	return content;
}
