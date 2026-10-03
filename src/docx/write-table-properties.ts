import type { Table, TableCell } from './model.js';
import type { TableBorders, TableCellMargins, TableRowProperties } from './table-model.js';
import { buildCellProperties, buildRowProperties } from './table-cell-write.js';
import { buildBorders } from './table-defaults.js';
import { patchMarginProperties } from './write-table-margins.js';
import { orderChildren } from './element-order.js';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';
import { isStVerticalJc } from './generated/wml-simple-types.js';

const ROW_ORDER = [
	'cnfStyle',
	'divId',
	'gridBefore',
	'gridAfter',
	'wBefore',
	'wAfter',
	'cantSplit',
	'trHeight',
	'tblHeader',
	'tblCellSpacing',
	'jc',
	'hidden',
	'ins',
	'del',
	'trPrChange',
];
const TABLE_ORDER = [
	'tblStyle',
	'tblpPr',
	'tblOverlap',
	'bidiVisual',
	'tblStyleRowBandSize',
	'tblStyleColBandSize',
	'tblW',
	'jc',
	'tblCellSpacing',
	'tblInd',
	'tblBorders',
	'shd',
	'tblLayout',
	'tblCellMar',
	'tblLook',
	'tblCaption',
	'tblDescription',
	'tblPrChange',
];

/** Changes only modeled row fields, preserving unmodeled trPr children and attributes. */
export function patchRowProperties(
	doc: XmlDocument,
	tr: XmlElement,
	next: TableRowProperties | undefined,
	base: TableRowProperties | undefined,
): void {
	if (JSON.stringify(next ?? {}) === JSON.stringify(base ?? {})) return;
	let props = first(tr, 'trPr');
	if (!props) {
		props = makeW(doc, 'trPr');
		tr.insertBefore(props, tr.firstChild);
	}
	const built = buildRowProperties(doc, next, 0);
	for (const [name, changed] of [
		['cantSplit', Boolean(next?.cantSplit) !== Boolean(base?.cantSplit)],
		['tblHeader', Boolean(next?.header) !== Boolean(base?.header)],
		['trHeight', next?.heightTwips !== base?.heightTwips || next?.heightRule !== base?.heightRule],
	] as const) {
		if (!changed) continue;
		for (const old of children(props, name)) props.removeChild(old);
		const replacement = built && first(built, name);
		if (replacement) props.appendChild(replacement);
	}
	orderChildren(props, ROW_ORDER);
}

/** Updates table defaults without replacing borders, style, look or extension XML. */
export function patchTableMargins(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base?: Table,
): void {
	patchMarginProperties(
		doc,
		node,
		'tblPr',
		'tblCellMar',
		TABLE_ORDER,
		table.cellMargins,
		base?.cellMargins,
	);
}

const CELL_ORDER = [
	'cnfStyle',
	'tcW',
	'gridSpan',
	'hMerge',
	'vMerge',
	'tcBorders',
	'shd',
	'noWrap',
	'tcMar',
	'textDirection',
	'tcFitText',
	'vAlign',
	'hideMark',
	'headers',
	'cellIns',
	'cellDel',
	'cellMerge',
	'tcPrChange',
];
const CELL_BORDER_ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV'];
const CELL_SIDES = ['top', 'left', 'bottom', 'right'] as const;

/** Patches only vertical alignment; undefined restores table/style default inheritance. */
export function patchCellVerticalAlignment(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableCell['verticalAlign'],
	base: TableCell['verticalAlign'],
): void {
	if (next === base) return;
	if (next !== undefined && !isStVerticalJc(next))
		throw new Error('Cell vertical alignment must be a valid ST_VerticalJc value.');
	let props = first(tc, 'tcPr');
	if (!props) {
		if (next === undefined) return;
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	const alignment = first(props, 'vAlign') ?? makeW(doc, 'vAlign');
	for (const old of children(props, 'vAlign')) props.removeChild(old);
	if (next !== undefined) {
		alignment.setAttributeNS(WORD_NS, 'w:val', next);
		props.appendChild(alignment);
	}
	orderChildren(props, CELL_ORDER);
	if (!Array.from(props.childNodes).some((node) => node.nodeType === 1) && !props.attributes.length)
		tc.removeChild(props);
}

/** Changes individual cell overrides; deleting an override restores table/default inheritance. */
export function patchCellMargins(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableCellMargins | undefined,
	base: TableCellMargins | undefined,
): void {
	patchMarginProperties(doc, tc, 'tcPr', 'tcMar', CELL_ORDER, next, base);
}

/**
 * Changes only the cell border sides that differ from `base`, keeping every other `tcPr` child
 * (width, shading, margins, extensions) and untouched sides exactly as they were.
 */
export function patchCellBorders(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableBorders | undefined,
	base: TableBorders | undefined,
): void {
	const changed = CELL_SIDES.filter(
		(side) => JSON.stringify(next?.[side]) !== JSON.stringify(base?.[side]),
	);
	if (!changed.length) return;
	let props = first(tc, 'tcPr');
	if (!props) {
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	let borders = first(props, 'tcBorders');
	if (!borders) {
		borders = makeW(doc, 'tcBorders');
		props.appendChild(borders);
	}
	const built = buildBorders(doc, next ?? {}, 'tcBorders');
	for (const side of changed) {
		for (const old of children(borders, side)) borders.removeChild(old);
		const replacement = first(built, side);
		if (replacement) borders.appendChild(replacement);
	}
	if (!Array.from(borders.childNodes).some((node) => node.nodeType === 1))
		props.removeChild(borders);
	else orderChildren(borders, CELL_BORDER_ORDER);
	orderChildren(props, CELL_ORDER);
	if (!Array.from(props.childNodes).some((node) => node.nodeType === 1)) tc.removeChild(props);
}

/** Replaces only changed cell shading, preserving widths, borders, margins and extension XML. */
export function patchCellShading(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableCell,
	base?: TableCell,
): void {
	if (
		next.shadingFill === base?.shadingFill &&
		JSON.stringify(next.shadingThemeFill) === JSON.stringify(base?.shadingThemeFill)
	)
		return;
	if (next.shadingFill && next.shadingFill !== 'auto' && !/^#?[0-9a-f]{6}$/i.test(next.shadingFill))
		throw new Error('Cell shading must be a six-digit RGB color or auto.');
	let props = first(tc, 'tcPr');
	if (!props) {
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	for (const old of children(props, 'shd')) props.removeChild(old);
	const built = buildCellProperties(doc, next);
	const replacement = built && first(built, 'shd');
	if (replacement) props.appendChild(replacement);
	orderChildren(props, CELL_ORDER);
	if (!Array.from(props.childNodes).some((node) => node.nodeType === 1)) tc.removeChild(props);
}
