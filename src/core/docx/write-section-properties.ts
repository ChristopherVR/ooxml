// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes the properties of one changed section into its w:sectPr: break type, page size and
// orientation, margins, columns, vertical alignment, first-page setting, line numbering, page
// borders and page numbering, each only when it changed, in CT_SectPr order.
import type { SectionProperties } from './model';
import { writeFootnoteColumns } from './write-footnote-columns';
import { patchPageBorders } from './page-borders';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

export function setW(element: XmlElement, name: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${name}`, value);
}

/** Returns the named child, creating it in the right place (schema order is honored by callers). */
function child(doc: XmlDocument, parent: XmlElement, name: string, before?: string[]): XmlElement {
	const existing = first(parent, name);
	if (existing) return existing;
	const element = makeW(doc, name);
	const anchor = before
		?.map((candidate) => first(parent, candidate))
		.find((candidate): candidate is XmlElement => Boolean(candidate));
	parent.insertBefore(element, anchor ?? null);
	return element;
}

// Elements that follow pgSz/pgMar/cols inside w:sectPr (CT_SectPr sequence).
const AFTER_PGSZ = [
	'pgMar',
	'paperSrc',
	'pgBorders',
	'lnNumType',
	'pgNumType',
	'cols',
	'formProt',
	'vAlign',
	'noEndnote',
	'titlePg',
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_COLS = [
	'formProt',
	'vAlign',
	'noEndnote',
	'titlePg',
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_PGNUMTYPE = ['cols', ...AFTER_COLS];
const AFTER_LNNUM = ['pgNumType', ...AFTER_PGNUMTYPE];
const AFTER_TITLEPG = [
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_PGBORDERS = ['lnNumType', ...AFTER_LNNUM];
const AFTER_TYPE = ['pgSz', ...AFTER_PGSZ];

export function writeSectionProperties(
	doc: XmlDocument,
	sectPr: XmlElement,
	section: SectionProperties,
	base?: SectionProperties,
) {
	if (section.footnoteColumns !== base?.footnoteColumns)
		writeFootnoteColumns(doc, sectPr, section.footnoteColumns);
	if (section.type === 'nextPage')
		for (const type of children(sectPr, 'type')) sectPr.removeChild(type);
	else setW(child(doc, sectPr, 'type', AFTER_TYPE), 'val', section.type);
	const size = child(doc, sectPr, 'pgSz', AFTER_PGSZ);
	setW(size, 'w', String(section.pageWidthTwips));
	setW(size, 'h', String(section.pageHeightTwips));
	if (section.orientation === 'landscape') setW(size, 'orient', 'landscape');
	else size.removeAttributeNS(WORD_NS, 'orient');
	const margins = child(doc, sectPr, 'pgMar', AFTER_PGSZ.slice(1));
	setW(margins, 'top', String(section.marginTopTwips));
	setW(margins, 'right', String(section.marginRightTwips));
	setW(margins, 'bottom', String(section.marginBottomTwips));
	setW(margins, 'left', String(section.marginLeftTwips));
	setW(margins, 'header', String(section.headerDistanceTwips ?? 720));
	setW(margins, 'footer', String(section.footerDistanceTwips ?? 720));
	setW(margins, 'gutter', String(section.gutterTwips ?? 0));
	const columns = child(doc, sectPr, 'cols', AFTER_COLS);
	setW(columns, 'num', String(section.columns.count));
	setW(columns, 'space', String(section.columns.spacingTwips ?? 720));
	if (section.columns.equalWidth || !section.columns.widths?.length) {
		for (const column of children(columns, 'col')) columns.removeChild(column);
		columns.removeAttributeNS(WORD_NS, 'equalWidth');
	} else {
		setW(columns, 'equalWidth', '0');
		const old = children(columns, 'col');
		section.columns.widths.forEach((value, index) => {
			const column = old[index] ?? makeW(doc, 'col');
			if (!old[index]) columns.appendChild(column);
			setW(column, 'w', String(value.widthTwips));
			if (value.spacingTwips !== undefined) setW(column, 'space', String(value.spacingTwips));
			else column.removeAttributeNS(WORD_NS, 'space');
		});
		for (const column of old.slice(section.columns.widths.length)) columns.removeChild(column);
	}
	if (section.columns.separator) setW(columns, 'sep', '1');
	else columns.removeAttributeNS(WORD_NS, 'sep');
	if (section.verticalAlign && section.verticalAlign !== 'top')
		setW(
			child(doc, sectPr, 'vAlign', ['noEndnote', 'titlePg', ...AFTER_TITLEPG]),
			'val',
			section.verticalAlign,
		);
	else for (const vAlign of children(sectPr, 'vAlign')) sectPr.removeChild(vAlign);
	if (section.titlePage)
		child(doc, sectPr, 'titlePg', AFTER_TITLEPG).removeAttributeNS(WORD_NS, 'val');
	else for (const titlePg of children(sectPr, 'titlePg')) sectPr.removeChild(titlePg);
	const line = section.lineNumberSettings;
	if (line) {
		const lnNum = child(doc, sectPr, 'lnNumType', AFTER_LNNUM);
		if (line.countBy > 1) setW(lnNum, 'countBy', String(line.countBy));
		else lnNum.removeAttributeNS(WORD_NS, 'countBy');
		if (line.start !== 1) setW(lnNum, 'start', String(line.start - 1));
		else lnNum.removeAttributeNS(WORD_NS, 'start');
		if (line.distanceTwips !== undefined) setW(lnNum, 'distance', String(line.distanceTwips));
		else lnNum.removeAttributeNS(WORD_NS, 'distance');
		if (line.restart !== 'continuous') setW(lnNum, 'restart', line.restart);
		else lnNum.removeAttributeNS(WORD_NS, 'restart');
	} else if (!section.lineNumbering)
		for (const lnNum of children(sectPr, 'lnNumType')) sectPr.removeChild(lnNum);
	patchPageBorders(doc, sectPr, section.pageBorders, base?.pageBorders, (d, parent, name) =>
		child(d, parent, name, AFTER_PGBORDERS),
	);
	const numbering = section.pageNumbering;
	if (numbering?.format || numbering?.start !== undefined) {
		const pgNumType = child(doc, sectPr, 'pgNumType', AFTER_PGNUMTYPE);
		if (numbering.format && numbering.format !== 'decimal')
			setW(pgNumType, 'fmt', numbering.format);
		else pgNumType.removeAttributeNS(WORD_NS, 'fmt');
		if (numbering.start !== undefined) setW(pgNumType, 'start', String(numbering.start));
		else pgNumType.removeAttributeNS(WORD_NS, 'start');
	} else
		for (const pgNumType of children(sectPr, 'pgNumType')) {
			pgNumType.removeAttributeNS(WORD_NS, 'fmt');
			pgNumType.removeAttributeNS(WORD_NS, 'start');
			if (!pgNumType.attributes.length) sectPr.removeChild(pgNumType);
		}
}
