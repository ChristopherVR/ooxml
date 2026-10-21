// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The content control values the writer writes as they are. A control's fields need not come from
// the document: the editor keeps an inline control in its span's `data-sdt` attribute, and a paste
// keeps another page's verbatim properties out (ui/content-control-sources.ts), but a model can be
// built by any host. So a control's modeled id, lock and kind are written only when Word reads
// them, and its verbatim `w:sdtPr` only when it is shaped as Word's schema has it: no content, and
// nothing out of place, ever reaches the saved document. Attribute values other than the id and
// the lock are not checked.
import { WORD_2012_NS } from './content-control';
import type { ContentControl } from './content-control-model';
import { RPR_ORDER } from './element-order';
import { WORD_2010_NS } from './ligatures';
import {
	children,
	elements,
	getW,
	parseXml,
	WORD_NS,
	type XmlDocument,
	type XmlElement,
} from './xml';

/** The lock, when it is one the model holds (`ST_Lock`'s locks; `unlocked` is the default). */
export const validLock = (lock: string | undefined): ContentControl['lock'] =>
	lock === 'sdtLocked' || lock === 'contentLocked' || lock === 'sdtContentLocked'
		? lock
		: undefined;

/** The id, when it is a decimal integer Word reads (signed 32-bit). */
export function validId(id: string | undefined): string | undefined {
	if (id === undefined || !/^-?\d{1,10}$/.test(id)) return undefined;
	const value = Number(id);
	return value >= -(2 ** 31) && value < 2 ** 31 ? id : undefined;
}

/** The kinds whose `w:sdtPr` type element is named after the kind; rich text has none. */
export const NAMED_TYPES: ReadonlySet<string> = new Set([
	'picture',
	'group',
	'equation',
	'citation',
	'bibliography',
	'docPartObj',
	'docPartList',
]);

/** `CT_SdtPr`'s type choices (ECMA-376 Part 1, 17.5.2.38): a control has one at most. */
const TYPE_CHOICES: ReadonlySet<string> = new Set([
	'equation',
	'comboBox',
	'date',
	'docPartObj',
	'docPartList',
	'dropDownList',
	'picture',
	'richText',
	'text',
	'citation',
	'group',
	'bibliography',
]);

/** `CT_SdtPr`'s elements: its properties, each once at most, and the type choices. */
const SDT_PROPERTIES: ReadonlySet<string> = new Set([
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
	...TYPE_CHOICES,
]);

/** What each element of a `w:sdtPr` holds (ECMA-376 Part 1, 17.5.2); the others hold none. */
const PROPERTY_CHILDREN: Readonly<Record<string, ReadonlySet<string>>> = {
	sdtPr: SDT_PROPERTIES,
	rPr: new Set(RPR_ORDER),
	rPrChange: new Set(['rPr']),
	placeholder: new Set(['docPart']),
	comboBox: new Set(['listItem']),
	dropDownList: new Set(['listItem']),
	date: new Set(['dateFormat', 'lid', 'storeMappedDataAs', 'calendar']),
	docPartObj: new Set(['docPartGallery', 'docPartCategory', 'docPartUnique']),
	docPartList: new Set(['docPartGallery', 'docPartCategory', 'docPartUnique']),
};

/** The namespaces of the extensions Word writes in a `w:sdtPr` or its `w:rPr` (`w14`, `w15`). */
const WORD_EXTENSIONS: ReadonlySet<string> = new Set([WORD_2010_NS, WORD_2012_NS]);

/**
 * The namespaces of the attributes a property may have: Word's, its extensions', those later
 * versions of Word add (a data binding's checksum, `w16sdtdh:storeItemChecksum`, [MS-DOCX]
 * 2.11.2.1, and a revision's UTC date, `w16du:dateUtc`), and XML's own.
 */
const ATTRIBUTE_NAMESPACES: ReadonlySet<string> = new Set([
	WORD_NS,
	...WORD_EXTENSIONS,
	'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
	'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
	'http://www.w3.org/XML/1998/namespace',
	'http://www.w3.org/2000/xmlns/',
]);

/**
 * The type elements, by namespace and name: `TYPE_CHOICES`, and those of Word's extensions (a check
 * box, an entity picker, a repeating section and its items).
 */
const TYPES: ReadonlySet<string> = new Set([
	...[...TYPE_CHOICES].map((local) => `${WORD_NS} ${local}`),
	`${WORD_2010_NS} checkbox`,
	`${WORD_2010_NS} entityPicker`,
	`${WORD_2012_NS} repeatingSection`,
	`${WORD_2012_NS} repeatingSectionItem`,
]);

/** The WordprocessingML elements Word puts its extensions in: `w14:checkbox`, `w14:glow`. */
const EXTENSIBLE: ReadonlySet<string> = new Set(['sdtPr', 'rPr']);

/**
 * Whether `element`, a `w:sdtPr` or one of its properties, holds only what its schema allows:
 * each WordprocessingML child one its element may hold, elements of Word's extension namespaces
 * where Word puts them, shaped the same way, attributes of those namespaces only (so no
 * relationship reference or markup compatibility directive), and no text.
 */
function propertyShaped(element: XmlElement): boolean {
	if (
		Array.from(element.attributes).some(
			(attribute) => !ATTRIBUTE_NAMESPACES.has(attribute.namespaceURI ?? ''),
		)
	)
		return false;
	const word = element.namespaceURI === WORD_NS;
	const allowed = word ? PROPERTY_CHILDREN[element.localName] : undefined;
	const extensible = !word || EXTENSIBLE.has(element.localName);
	return Array.from(element.childNodes).every((child) => {
		if (child.nodeType === 3 || child.nodeType === 4) return !child.textContent?.trim();
		if (child.nodeType !== 1) return true;
		const item = child as XmlElement;
		const known =
			item.namespaceURI === WORD_NS
				? allowed?.has(item.localName) === true
				: extensible && WORD_EXTENSIONS.has(item.namespaceURI ?? '');
		return known && propertyShaped(item);
	});
}

/**
 * Whether a `w:sdtPr` holds only properties Word reads: shaped as its schema allows, each property
 * once at most, Word's extensions included, one type at most, and an id and a lock whose values
 * Word reads.
 */
function onlyProperties(sdtPr: XmlElement): boolean {
	if (!propertyShaped(sdtPr)) return false;
	const names = elements(sdtPr).map((child) => `${child.namespaceURI} ${child.localName}`);
	if (new Set(names).size < names.length) return false;
	if (names.filter((name) => TYPES.has(name)).length > 1) return false;
	const value = (local: string) => children(sdtPr, local).map((child) => getW(child, 'val'));
	return (
		value('id').every((id) => validId(id) !== undefined) &&
		value('lock').every((lock) => lock === 'unlocked' || validLock(lock) !== undefined)
	);
}

/**
 * The `w:sdtPr` that `xml` holds, imported into `doc`: nothing when it is not XML, holds another
 * element, or holds anything but properties Word reads, such as a paragraph with a field that
 * fetches a picture when it is updated, or an id or lock Word does not read.
 */
export function verbatimProperties(
	doc: XmlDocument,
	xml: string | undefined,
): XmlElement | undefined {
	if (!xml) return undefined;
	let root: XmlElement | null;
	try {
		root = parseXml(xml).documentElement;
	} catch {
		return undefined;
	}
	return root?.localName === 'sdtPr' && root.namespaceURI === WORD_NS && onlyProperties(root)
		? (doc.importNode(root, true) as XmlElement)
		: undefined;
}
