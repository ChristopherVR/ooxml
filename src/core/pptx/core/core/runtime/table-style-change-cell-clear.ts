import type { PptxTableCellStyle, XmlObject } from '../../types';

/**
 * Saving a table whose style changed: drop the direct cell fills and borders
 * the model no longer carries.
 *
 * PowerPoint's table-style pick clears every cell's direct `a:tcPr` fill and
 * borders and keeps the run formatting (`Table.ApplyStyle`, checked over COM;
 * see `ooxml-ui/pptx`'s `table-styles-gallery.ts`). The editor does the same
 * in memory, but `writeTableCellStyle` only ever writes a fill or a border
 * that is set, so the cell's original children stayed in the file and the
 * old look came back on reload.
 *
 * It runs when the saved `a:tableStyleId` differs from the file's, or when
 * the table data carries `clearDirectCellFormatting` (set by the pick, so the
 * same style picked again, or B then back to A, is covered too).
 *
 * Only children the model has no value for are removed, so a fill or border
 * set after the pick (or kept by a caller that changes the id without
 * clearing, such as the inspector's style select or the `assignTableStyle`
 * tool) is still written by `writeTableCellStyle` as before. A direct child
 * the parser cannot model (a line that is only `a:noFill`, a line with a
 * gradient fill and no width, an unresolved `a:blipFill`, an `a:grpFill`) has
 * no model value either, so a style change removes it on every path.
 */

/** `EG_FillProperties` children of `a:tcPr`. */
const FILL_KEYS = [
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
];

/** Each `a:tcPr` line child and the model fields that describe it. */
const LINE_KEYS: ReadonlyArray<readonly [string, ReadonlyArray<keyof PptxTableCellStyle>]> = [
	['a:lnL', ['borderLeftWidth', 'borderLeftColor', 'borderLeftDash']],
	['a:lnR', ['borderRightWidth', 'borderRightColor', 'borderRightDash']],
	['a:lnT', ['borderTopWidth', 'borderTopColor', 'borderTopDash']],
	['a:lnB', ['borderBottomWidth', 'borderBottomColor', 'borderBottomDash']],
	['a:lnTlToBr', ['borderDiagDownWidth', 'borderDiagDownColor']],
	['a:lnBlToTr', ['borderDiagUpWidth', 'borderDiagUpColor']],
];

/** The model fields that mean the cell keeps a direct fill of its own. */
const FILL_FIELDS: ReadonlyArray<keyof PptxTableCellStyle> = [
	'fillMode',
	'backgroundColor',
	'gradientFillStops',
	'patternFillPreset',
	'backgroundImageFillPath',
	'backgroundImageFillData',
];

/** Normalise a style GUID for comparison: trimmed, upper case, braced. */
function normaliseStyleId(id: unknown): string {
	const text = typeof id === 'string' || typeof id === 'number' ? String(id) : '';
	const trimmed = text.trim().toUpperCase();
	if (!trimmed) {
		return '';
	}
	return trimmed.startsWith('{') ? trimmed : `{${trimmed}}`;
}

/**
 * The style id a parsed `a:tblPr` carries, or '': the spec `a:tableStyleId`
 * text first, then the legacy `a:tblStyle@val` and `@tblStyle` forms the
 * parser also accepts (`PptxTableDataParser.extractTableStyleId`), so a file
 * in a legacy form does not count as a style change on every save.
 */
function xmlTableStyleId(tblPr: unknown): string {
	if (!tblPr || typeof tblPr !== 'object') {
		return '';
	}
	const props = tblPr as XmlObject;
	const node = props['a:tableStyleId'];
	const spec = normaliseStyleId(
		node && typeof node === 'object' ? (node as XmlObject)['#text'] : node,
	);
	if (spec) {
		return spec;
	}
	const legacy = props['a:tblStyle'] as XmlObject | undefined;
	return normaliseStyleId(legacy?.['@_val'] || props['@_tblStyle']);
}

/**
 * Whether saving `nextStyleId` onto a table whose `a:tblPr` is `tblPr`
 * changes its style. `undefined` keeps the file's id (see
 * `serializeTablePropertyFlags`), so it is never a change.
 */
export function tableStyleIdChanged(tblPr: unknown, nextStyleId: string | undefined): boolean {
	if (nextStyleId === undefined) {
		return false;
	}
	return normaliseStyleId(nextStyleId) !== xmlTableStyleId(tblPr);
}

const hasAny = (
	style: PptxTableCellStyle | undefined,
	keys: ReadonlyArray<keyof PptxTableCellStyle>,
) => keys.some((key) => style?.[key] !== undefined);

/**
 * Remove the direct fill, borders and `a:cell3D` from a cell's `a:tcPr` when
 * the cell's model `style` has no value for them.
 */
export function clearCellFormattingForStyleChange(
	xmlCell: XmlObject,
	style: PptxTableCellStyle | undefined,
): void {
	const tcPr = xmlCell['a:tcPr'];
	if (!tcPr || typeof tcPr !== 'object') {
		return;
	}
	const props = tcPr as XmlObject;
	if (!hasAny(style, FILL_FIELDS)) {
		for (const key of FILL_KEYS) {
			delete props[key];
		}
	}
	for (const [key, fields] of LINE_KEYS) {
		if (!hasAny(style, fields)) {
			delete props[key];
		}
	}
	if (!style?.cell3D) {
		delete props['a:cell3D'];
	}
}
