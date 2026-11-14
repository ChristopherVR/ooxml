/**
 * Apply a table cell's text style updates to one run's `a:rPr` (or an
 * `a:endParaRPr`), for the raw-XML half of cell formatting.
 *
 * @module runtime/table-cell-run-style-xml
 */
import type { XmlObject } from '../../types';

/** Set a run attribute to `on` when `value` is truthy, remove it otherwise. */
function setRunFlag(rPr: XmlObject, attribute: string, value: unknown, on: string): void {
	if (value) {
		rPr[attribute] = on;
	} else {
		delete rPr[attribute];
	}
}

/** CT_TextCharacterProperties children that follow `a:latin`. */
const AFTER_LATIN = [
	'a:ea',
	'a:cs',
	'a:sym',
	'a:hlinkClick',
	'a:hlinkMouseOver',
	'a:rtl',
	'a:extLst',
] as const;

/** Set `a:latin@typeface`, keeping the schema order of the children after it. */
function setRunLatinTypeface(rPr: XmlObject, typeface: string): void {
	const existing = rPr['a:latin'];
	const latin: XmlObject = {
		...(existing && typeof existing === 'object' ? (existing as XmlObject) : {}),
		'@_typeface': typeface,
	};
	const tail: Array<[string, XmlObject[string]]> = [];
	for (const key of AFTER_LATIN) {
		if (key in rPr) {
			tail.push([key, rPr[key]]);
			delete rPr[key];
		}
	}
	delete rPr['a:latin'];
	rPr['a:latin'] = latin;
	for (const [key, value] of tail) {
		rPr[key] = value;
	}
}

/**
 * Apply bold, italic, underline, strikethrough, size (points, as on
 * `PptxTableCellStyle`), colour and font family updates to a run property
 * node in place. Keys absent from `styleUpdates` are left alone.
 */
export function applyCellRunStyleUpdates(
	rPr: XmlObject,
	styleUpdates: Record<string, unknown>,
): void {
	if ('bold' in styleUpdates) {
		setRunFlag(rPr, '@_b', styleUpdates.bold, '1');
	}
	if ('italic' in styleUpdates) {
		setRunFlag(rPr, '@_i', styleUpdates.italic, '1');
	}
	if ('underline' in styleUpdates) {
		setRunFlag(rPr, '@_u', styleUpdates.underline, 'sng');
	}
	if ('strikethrough' in styleUpdates) {
		setRunFlag(rPr, '@_strike', styleUpdates.strikethrough, 'sngStrike');
	}
	if (typeof styleUpdates.fontSize === 'number' && styleUpdates.fontSize > 0) {
		rPr['@_sz'] = String(Math.round(styleUpdates.fontSize * 100));
	}
	if ('color' in styleUpdates && typeof styleUpdates.color === 'string') {
		const hex = styleUpdates.color.replace('#', '');
		const tail: Array<[string, XmlObject[string]]> = [];
		// The fill precedes the effects, highlight and fonts (CT_TextCharacterProperties).
		for (const key of Object.keys(rPr)) {
			if (key.startsWith('a:') && !['a:ln'].includes(key)) {
				tail.push([key, rPr[key]]);
				delete rPr[key];
			}
		}
		rPr['a:solidFill'] = { 'a:srgbClr': { '@_val': hex } };
		for (const [key, value] of tail) {
			if (
				![
					'a:solidFill',
					'a:noFill',
					'a:gradFill',
					'a:pattFill',
					'a:blipFill',
					'a:grpFill',
				].includes(key)
			) {
				rPr[key] = value;
			}
		}
	}
	if (typeof styleUpdates.fontFamily === 'string' && styleUpdates.fontFamily) {
		setRunLatinTypeface(rPr, styleUpdates.fontFamily);
	}
}
