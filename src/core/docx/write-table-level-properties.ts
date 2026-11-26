// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Source-preserving XML patchers for table-level properties (`w:tblPr`, `w:tblGrid`) and cell
// width (`w:tcW`): only fields that changed relative to `base` are updated, keeping every other
// child element, attribute and extension XML untouched.
import type { Table, TableCell } from './model';
import type { TableBorders } from './table-model';
import {
	buildBorders,
	buildTableLook,
	TABLE_LOOK_FLAGS,
	tableJustification,
} from './table-defaults';
import { CELL_BORDER_ORDER, CELL_ORDER, TABLE_ORDER } from './write-table-properties';
import { orderChildren } from './element-order';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

const TABLE_BORDER_SIDES = [
	'top',
	'left',
	'bottom',
	'right',
	'insideH',
	'insideV',
] as const satisfies readonly (keyof TableBorders)[];

function setW(element: XmlElement, name: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${name}`, value);
}

function replaceChild(props: XmlElement, name: string, replacement: XmlElement | undefined): void {
	for (const old of children(props, name)) props.removeChild(old);
	if (replacement) props.appendChild(replacement);
}

function patchTableBorders(
	doc: XmlDocument,
	props: XmlElement,
	next: TableBorders | undefined,
	changed: readonly (keyof TableBorders)[],
): void {
	let borders = first(props, 'tblBorders');
	if (!borders) {
		borders = makeW(doc, 'tblBorders');
		props.appendChild(borders);
	}
	const built = buildBorders(doc, next ?? {}, 'tblBorders');
	for (const side of changed) {
		for (const old of children(borders, side)) borders.removeChild(old);
		if (side === 'left') for (const old of children(borders, 'start')) borders.removeChild(old);
		if (side === 'right') for (const old of children(borders, 'end')) borders.removeChild(old);
		const replacement = first(built, side);
		if (replacement) borders.appendChild(replacement);
	}
	if (!Array.from(borders.childNodes).some((node) => node.nodeType === 1))
		props.removeChild(borders);
	else orderChildren(borders, CELL_BORDER_ORDER);
}

/**
 * Patches changed table-level properties (`tblStyle`, `tblW`, `jc`, `tblInd`, `tblBorders`,
 * `tblLook`) in `w:tblPr`, preserving every untouched property and unmodeled extension XML.
 */
export function patchTableProperties(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base?: Table,
): void {
	const styleChanged = table.style !== base?.style;
	const widthChanged =
		table.widthTwips !== base?.widthTwips || table.widthPercent !== base?.widthPercent;
	const jcChanged =
		table.alignment !== base?.alignment || table.justification !== base?.justification;
	const indentChanged = table.indentTwips !== base?.indentTwips;
	const bordersChanged = TABLE_BORDER_SIDES.filter(
		(side) => JSON.stringify(table.borders?.[side]) !== JSON.stringify(base?.borders?.[side]),
	);
	const lookChanged = JSON.stringify(table.look) !== JSON.stringify(base?.look);
	const layoutChanged = table.layout !== base?.layout;
	if (
		!styleChanged &&
		!widthChanged &&
		!jcChanged &&
		!indentChanged &&
		!bordersChanged.length &&
		!lookChanged &&
		!layoutChanged
	)
		return;
	let props = first(node, 'tblPr');
	if (!props) {
		props = makeW(doc, 'tblPr');
		node.insertBefore(props, node.firstChild);
	}
	if (styleChanged) {
		const style = table.style ? makeW(doc, 'tblStyle') : undefined;
		if (style && table.style) setW(style, 'val', table.style);
		replaceChild(props, 'tblStyle', style);
	}
	if (widthChanged) {
		const width = makeW(doc, 'tblW');
		if (table.widthTwips !== undefined) {
			setW(width, 'w', String(table.widthTwips));
			setW(width, 'type', 'dxa');
		} else if (table.widthPercent !== undefined) {
			setW(width, 'w', String(Math.round(table.widthPercent * 5000)));
			setW(width, 'type', 'pct');
		} else {
			setW(width, 'w', '0');
			setW(width, 'type', 'auto');
		}
		replaceChild(props, 'tblW', width);
	}
	if (jcChanged) {
		const jcValue = tableJustification(table);
		const jc = jcValue ? makeW(doc, 'jc') : undefined;
		if (jc && jcValue) setW(jc, 'val', jcValue);
		replaceChild(props, 'jc', jc);
	}
	if (indentChanged) {
		const indent = table.indentTwips !== undefined ? makeW(doc, 'tblInd') : undefined;
		if (indent && table.indentTwips !== undefined) {
			setW(indent, 'w', String(table.indentTwips));
			setW(indent, 'type', 'dxa');
		}
		replaceChild(props, 'tblInd', indent);
	}
	if (bordersChanged.length) patchTableBorders(doc, props, table.borders, bordersChanged);
	if (layoutChanged) {
		const layout = table.layout ? makeW(doc, 'tblLayout') : undefined;
		if (layout && table.layout) setW(layout, 'type', table.layout);
		replaceChild(props, 'tblLayout', layout);
	}
	if (lookChanged) {
		const hasLook =
			table.look && TABLE_LOOK_FLAGS.some(([name]) => table.look?.[name] !== undefined);
		replaceChild(props, 'tblLook', hasLook ? buildTableLook(doc, table.look!) : undefined);
	}
	orderChildren(props, TABLE_ORDER);
}

/**
 * Brings `w:tblGrid` to `table.grid` when `table.grid` changed relative to `base`, preserving
 * surviving `w:gridCol` elements and creating `w:tblGrid` after `w:tblPr` when absent.
 */
export function patchTableGrid(
	doc: XmlDocument,
	node: XmlElement,
	table: Table,
	base?: Table,
): void {
	if (!table.grid?.length) return;
	if (base && JSON.stringify(table.grid) === JSON.stringify(base.grid)) return;
	let grid = first(node, 'tblGrid');
	if (!grid) {
		grid = makeW(doc, 'tblGrid');
		const tblPr = first(node, 'tblPr');
		node.insertBefore(grid, tblPr ? tblPr.nextSibling : node.firstChild);
	}
	const old = children(grid, 'gridCol');
	table.grid.forEach((width, index) => {
		let column = old[index];
		if (!column) {
			column = makeW(doc, 'gridCol');
			grid.appendChild(column);
		}
		if (column.getAttributeNS(WORD_NS, 'w') !== String(width)) setW(column, 'w', String(width));
	});
	for (const extra of old.slice(table.grid.length)) grid.removeChild(extra);
}

/** Patches only `w:tcW` on a surviving cell, keeping every other `w:tcPr` child intact. */
export function patchCellWidth(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableCell,
	base: TableCell,
): void {
	if (next.widthTwips === base.widthTwips && next.widthPercent === base.widthPercent) return;
	const width = next.widthTwips ?? next.widthPercent;
	let props = first(tc, 'tcPr');
	if (!props) {
		if (width === undefined) return;
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	if (width === undefined) replaceChild(props, 'tcW', undefined);
	else {
		const old = first(props, 'tcW');
		const element = old ?? makeW(doc, 'tcW');
		setW(element, 'w', String(next.widthTwips === undefined ? Math.round(width * 5000) : width));
		setW(element, 'type', next.widthTwips === undefined ? 'pct' : 'dxa');
		if (!old) props.appendChild(element);
	}
	orderChildren(props, CELL_ORDER);
	if (
		!Array.from(props.childNodes).some((child) => child.nodeType === 1) &&
		!props.attributes.length
	)
		tc.removeChild(props);
}
