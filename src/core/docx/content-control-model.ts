// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The model of a structured document tag (`w:sdt`, a Word content control): its kind, identity,
// lock settings and the state each kind edits. Anything not modeled stays in `sourceXml`.

export type ContentControlKind =
	| 'richText'
	| 'text'
	| 'checkbox'
	| 'dropDownList'
	| 'comboBox'
	| 'date'
	| 'picture'
	| 'group'
	| 'equation'
	| 'citation'
	| 'bibliography'
	| 'docPartObj'
	| 'docPartList'
	| 'repeatingSection'
	| 'unknown';

/** Where the `w:sdt` sits: around runs, around paragraphs/tables, around a row, or around a cell. */
export type ContentControlLevel = 'inline' | 'block' | 'row' | 'cell';

/** `w:lock`: `sdtLocked` cannot be deleted, `contentLocked` cannot be edited, both together. */
export type ContentControlLock = 'sdtLocked' | 'contentLocked' | 'sdtContentLocked';

export interface ContentControlListItem {
	display: string;
	value: string;
}

/** A checkbox state glyph (`w14:checkedState`/`w14:uncheckedState`): a hex code point and font. */
export interface ContentControlGlyph {
	/** Hexadecimal code point, e.g. `2612`. */
	code: string;
	font?: string;
}

export interface ContentControl {
	/** `w:sdtPr/w:id`, or a generated id when the file has none. Shared by every run in the control. */
	id: string;
	level: ContentControlLevel;
	kind: ContentControlKind;
	/** `w:alias`, the title Word shows on the control's tab. */
	alias?: string;
	/** `w:tag`. */
	tag?: string;
	lock?: ContentControlLock;
	/** `w:showingPlcHdr`: the content is the placeholder, not a value. */
	showingPlaceholder?: boolean;
	/** `w:text/@w:multiLine`. */
	multiLine?: boolean;
	/** `w14:checkbox/w14:checked`. */
	checked?: boolean;
	checkedGlyph?: ContentControlGlyph;
	uncheckedGlyph?: ContentControlGlyph;
	/** `w:dropDownList`/`w:comboBox` `w:listItem` entries, in order. */
	items?: ContentControlListItem[];
	/** `w:date/w:dateFormat`, a Word date picture such as `dd MMMM yyyy`. */
	dateFormat?: string;
	/** `w:date/w:lid`, a language tag such as `en-US`. */
	dateLocale?: string;
	/** `w:date/@w:fullDate`, an ISO date-time. */
	fullDate?: string;
	/** The control carries `w:dataBinding`: kept verbatim, its custom XML part is not synchronized. */
	dataBound?: boolean;
	/** `w:temporary`: Word removes the control once the user edits it. */
	temporary?: boolean;
	/** `w:placeholder/w:docPart`: the glossary entry that holds the control's placeholder text. */
	placeholder?: string;
	/** The text of that glossary entry, which the control shows while it is empty (read only). */
	placeholderText?: string;
	/** Verbatim `w:sdtPr`, used to recreate the control and to keep what is not modeled. */
	sourceXml?: string;
	/** Verbatim `w:sdtEndPr` (the formatting of the control's end), kept when it is recreated. */
	endSourceXml?: string;
}

/** A content control surrounding a row or cell is read and preserved; the model keeps its fields. */
export const isContentControlLocked = (
	control: ContentControl | undefined,
	scope: 'content' | 'control',
): boolean => {
	if (!control?.lock) return false;
	return scope === 'content'
		? control.lock === 'contentLocked' || control.lock === 'sdtContentLocked'
		: control.lock === 'sdtLocked' || control.lock === 'sdtContentLocked';
};
