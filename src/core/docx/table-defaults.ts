// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Properties for tables created by the editor: Word requires `w:tblPr` and `w:tblGrid` on every
// table, and "Insert Table" gives new tables single-line borders on every edge.
import { fractionToThemeByte } from './theme-color';
import type { Table } from './model';
import type { StJcTable } from './generated/wml-simple-types';
import type { TableBorderSide, TableBorders, TableCellMargins, TableLook } from './table-model';
import { eighthPoints } from './units';
import { makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

export const TABLE_LOOK_FLAGS = [
	['firstRow', 0x0020],
	['lastRow', 0x0040],
	['firstColumn', 0x0080],
	['lastColumn', 0x0100],
	['noHBand', 0x0200],
	['noVBand', 0x0400],
] as const;

export const DEFAULT_TABLE_LOOK: TableLook = {
	firstRow: true,
	lastRow: false,
	firstColumn: true,
	lastColumn: false,
	noHBand: false,
	noVBand: true,
};

const SINGLE: TableBorderSide = { style: 'single', sizeEighthPoints: eighthPoints(4) };

/** Word's default borders for a newly inserted table (single 1/2 pt lines, automatic color). */
export const DEFAULT_TABLE_BORDERS: TableBorders = {
	top: SINGLE,
	left: SINGLE,
	bottom: SINGLE,
	right: SINGLE,
	insideH: SINGLE,
	insideV: SINGLE,
};

const BORDER_ORDER = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;

function setW(element: XmlElement, name: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${name}`, value);
}

/** Serializes `borders` as `w:tblBorders` (or `w:tcBorders`) in schema order. */
export function buildBorders(
	doc: XmlDocument,
	borders: TableBorders,
	tag: 'tblBorders' | 'tcBorders',
): XmlElement {
	const element = makeW(doc, tag);
	for (const side of BORDER_ORDER) {
		const value = borders[side];
		if (!value) continue;
		const edge = makeW(doc, side);
		setW(edge, 'val', value.style ?? 'single');
		setW(edge, 'sz', String(value.sizeEighthPoints ?? 4));
		setW(edge, 'space', '0');
		setW(edge, 'color', value.color ? value.color.replace(/^#/, '').toUpperCase() : 'auto');
		if (value.themeColor) {
			setW(edge, 'themeColor', value.themeColor);
			if (value.themeTint !== undefined)
				setW(edge, 'themeTint', fractionToThemeByte(value.themeTint));
			if (value.themeShade !== undefined)
				setW(edge, 'themeShade', fractionToThemeByte(value.themeShade));
		}
		element.appendChild(edge);
	}
	return element;
}

/** `w:tblCellMar` / `w:tcMar` with children in schema order (top, left, bottom, right). */
export function buildMargins(
	doc: XmlDocument,
	margins: TableCellMargins,
	tag: 'tblCellMar' | 'tcMar',
): XmlElement | undefined {
	const element = makeW(doc, tag);
	for (const side of ['top', 'left', 'bottom', 'right'] as const) {
		const value = margins[side];
		if (value === undefined) continue;
		const edge = makeW(doc, side);
		setW(edge, 'w', String(value));
		setW(edge, 'type', 'dxa');
		element.appendChild(edge);
	}
	return element.childNodes.length ? element : undefined;
}

/** The `w:tblPr/w:jc` value: the exact `justification` while `alignment` still matches it, else `alignment`. */
export function tableJustification(table: Table): StJcTable | undefined {
	const { alignment, justification } = table;
	const implied =
		justification === 'start' ? 'left' : justification === 'end' ? 'right' : justification;
	return justification && implied === alignment ? justification : alignment;
}

/** Builds a `w:tblLook` element with both the packed `w:val` bitmask and explicit flag attributes. */
export function buildTableLook(doc: XmlDocument, look: TableLook): XmlElement {
	let mask = 0;
	for (const [name, bit] of TABLE_LOOK_FLAGS) if (look[name]) mask |= bit;
	const element = makeW(doc, 'tblLook');
	setW(element, 'val', mask.toString(16).toUpperCase().padStart(4, '0'));
	for (const [name] of TABLE_LOOK_FLAGS) {
		const value = look[name];
		if (value !== undefined) setW(element, name, value ? '1' : '0');
	}
	return element;
}

/** Number of grid columns a table spans (the widest row, counting horizontal merges). */
export function tableColumnCount(table: Table): number {
	return Math.max(
		1,
		...table.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
	);
}

/** Grid column widths for a table the editor created: its grid, else the content width shared. */
export function newTableGridWidths(table: Table, contentWidthTwips: number): number[] {
	const columns = tableColumnCount(table);
	return table.grid?.length === columns
		? table.grid
		: Array.from({ length: columns }, () => Math.max(1, Math.floor(contentWidthTwips / columns)));
}

/**
 * Builds `w:tblPr` and `w:tblGrid` for a table the editor created. Column widths come from the
 * model's grid when present, otherwise the content width is shared equally.
 */
export function buildNewTableProperties(
	doc: XmlDocument,
	table: Table,
	contentWidthTwips: number,
): { tblPr: XmlElement; tblGrid: XmlElement } {
	const tblPr = makeW(doc, 'tblPr');
	if (table.style) {
		const style = makeW(doc, 'tblStyle');
		setW(style, 'val', table.style);
		tblPr.appendChild(style);
	}
	const width = makeW(doc, 'tblW');
	const percent = !table.widthTwips && table.widthPercent ? table.widthPercent : undefined;
	setW(
		width,
		'w',
		String(percent !== undefined ? Math.round(percent * 5000) : (table.widthTwips ?? 0)),
	);
	setW(width, 'type', table.widthTwips ? 'dxa' : percent !== undefined ? 'pct' : 'auto');
	tblPr.appendChild(width);
	const jcValue = tableJustification(table);
	if (jcValue) {
		const jc = makeW(doc, 'jc');
		setW(jc, 'val', jcValue);
		tblPr.appendChild(jc);
	}
	if (table.indentTwips !== undefined) {
		const indent = makeW(doc, 'tblInd');
		setW(indent, 'w', String(table.indentTwips));
		setW(indent, 'type', 'dxa');
		tblPr.appendChild(indent);
	}
	const borders = table.borders && buildBorders(doc, table.borders, 'tblBorders');
	if (borders?.childNodes.length) tblPr.appendChild(borders);
	if (table.layout) {
		const layout = makeW(doc, 'tblLayout');
		setW(layout, 'type', table.layout);
		tblPr.appendChild(layout);
	}
	const margins = table.cellMargins && buildMargins(doc, table.cellMargins, 'tblCellMar');
	if (margins) tblPr.appendChild(margins);
	const hasLook = table.look && TABLE_LOOK_FLAGS.some(([name]) => table.look?.[name] !== undefined);
	tblPr.appendChild(buildTableLook(doc, hasLook ? table.look! : DEFAULT_TABLE_LOOK));
	const tblGrid = makeW(doc, 'tblGrid');
	for (const value of newTableGridWidths(table, contentWidthTwips)) {
		const column = makeW(doc, 'gridCol');
		setW(column, 'w', String(value));
		tblGrid.appendChild(column);
	}
	return { tblPr, tblGrid };
}
