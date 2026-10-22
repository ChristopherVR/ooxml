// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads the properties of a `w:sdt` into the content control model.
import { children as nsChildren, first as nsFirst } from '../xml/index';
import type {
	ContentControl,
	ContentControlGlyph,
	ContentControlKind,
	ContentControlLevel,
	ContentControlLock,
} from './content-control-model';
import { parseOnOff } from './simple-types';
import { buildXml, first, getW, isElement, named, WORD_NS, type XmlElement } from './xml';

import { WORD_2010_NS } from './ligatures';
export const WORD_2012_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** The direct `w:sdtContent` of a `w:sdt`. */
export const sdtContent = (sdt: XmlElement): XmlElement | undefined => first(sdt, 'sdtContent');
export const isSdt = (element: XmlElement | null | undefined): element is XmlElement =>
	named(element, 'sdt');

const TYPE_ELEMENTS: Record<string, ContentControlKind> = {
	text: 'text',
	dropDownList: 'dropDownList',
	comboBox: 'comboBox',
	date: 'date',
	picture: 'picture',
	group: 'group',
	equation: 'equation',
	citation: 'citation',
	bibliography: 'bibliography',
	docPartObj: 'docPartObj',
	docPartList: 'docPartList',
};

function attrW14(element: XmlElement | undefined, local: string): string | undefined {
	return element?.getAttributeNS(WORD_2010_NS, local) ?? undefined;
}

function glyph(element: XmlElement | undefined): ContentControlGlyph | undefined {
	const code = attrW14(element, 'val');
	if (!code) return undefined;
	const font = attrW14(element, 'font');
	return { code, ...(font ? { font } : {}) };
}

/** The kind of control a child of `w:sdtPr` makes it, if it is a type element. */
function kindOf(child: XmlElement): ContentControlKind | undefined {
	if (child.namespaceURI === WORD_NS)
		return Object.hasOwn(TYPE_ELEMENTS, child.localName)
			? TYPE_ELEMENTS[child.localName]
			: undefined;
	if (child.namespaceURI === WORD_2010_NS && child.localName === 'checkbox') return 'checkbox';
	if (child.namespaceURI === WORD_2012_NS && child.localName === 'repeatingSection')
		return 'repeatingSection';
	return undefined;
}

/**
 * The type elements of a `w:sdtPr`, one at most in a valid file: also `w:richText`, which the
 * default type need not be written as, an entity picker and a repeating section's item.
 */
export const typeElements = (sdtPr: XmlElement): XmlElement[] =>
	Array.from(sdtPr.childNodes)
		.filter(isElement)
		.filter(
			(child) =>
				kindOf(child) !== undefined ||
				(child.namespaceURI === WORD_NS && child.localName === 'richText') ||
				(child.namespaceURI === WORD_2010_NS && child.localName === 'entityPicker') ||
				(child.namespaceURI === WORD_2012_NS && child.localName === 'repeatingSectionItem'),
		);

export function parseKind(sdtPr: XmlElement): ContentControlKind {
	const [type] = typeElements(sdtPr);
	// `w:richText` is the default type when no other type element is present.
	return (type && kindOf(type)) ?? 'richText';
}

let generated = 0;

/** Reads `w:sdtPr` into the control model. `level` is where the `w:sdt` sits. */
export function parseContentControl(sdt: XmlElement, level: ContentControlLevel): ContentControl {
	const sdtPr = first(sdt, 'sdtPr');
	if (!sdtPr) return { id: `sdt-gen-${++generated}`, level, kind: 'richText' };
	const kind = parseKind(sdtPr);
	const control: ContentControl = {
		id: getW(first(sdtPr, 'id'), 'val') ?? `sdt-gen-${++generated}`,
		level,
		kind,
	};
	const alias = getW(first(sdtPr, 'alias'), 'val');
	if (alias !== undefined) control.alias = alias;
	const tag = getW(first(sdtPr, 'tag'), 'val');
	if (tag !== undefined) control.tag = tag;
	const lock = getW(first(sdtPr, 'lock'), 'val');
	if (lock === 'sdtLocked' || lock === 'contentLocked' || lock === 'sdtContentLocked')
		control.lock = lock as ContentControlLock;
	const showing = first(sdtPr, 'showingPlcHdr');
	if (showing) control.showingPlaceholder = parseOnOff(getW(showing, 'val')) ?? true;
	if (first(sdtPr, 'dataBinding')) control.dataBound = true;
	const temporary = first(sdtPr, 'temporary');
	if (temporary && (parseOnOff(getW(temporary, 'val')) ?? true)) control.temporary = true;
	const placeholder = getW(first(first(sdtPr, 'placeholder'), 'docPart'), 'val');
	if (placeholder) control.placeholder = placeholder;
	if (kind === 'text') {
		const multi = parseOnOff(getW(first(sdtPr, 'text'), 'multiLine'));
		if (multi) control.multiLine = true;
	}
	if (kind === 'checkbox') {
		const box = nsChildren(sdtPr, 'checkbox', WORD_2010_NS, false)[0];
		const checked = nsFirst(box, 'checked', WORD_2010_NS, false);
		control.checked = parseOnOff(attrW14(checked, 'val')) ?? false;
		const on = glyph(nsFirst(box, 'checkedState', WORD_2010_NS, false));
		const off = glyph(nsFirst(box, 'uncheckedState', WORD_2010_NS, false));
		if (on) control.checkedGlyph = on;
		if (off) control.uncheckedGlyph = off;
	}
	if (kind === 'dropDownList' || kind === 'comboBox') {
		const list = first(sdtPr, kind);
		control.items = (list ? Array.from(list.childNodes).filter(isElement) : [])
			.filter((item) => named(item, 'listItem'))
			.map((item) => {
				const display = getW(item, 'displayText');
				const value = getW(item, 'value') ?? display ?? '';
				return { display: display ?? value, value };
			});
	}
	if (kind === 'date') {
		const date = first(sdtPr, 'date');
		const fullDate = getW(date, 'fullDate');
		if (fullDate) control.fullDate = fullDate;
		const format = getW(first(date, 'dateFormat'), 'val');
		if (format) control.dateFormat = format;
		const lid = getW(first(date, 'lid'), 'val');
		if (lid) control.dateLocale = lid;
	}
	control.sourceXml = buildXml(sdtPr);
	const end = first(sdt, 'sdtEndPr');
	if (end) control.endSourceXml = buildXml(end);
	return control;
}
