// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	NestedTablePreview,
	TableCell,
	TableCellMargins,
	TableLook,
	TableRowProperties,
} from './table-model';
import type { Paragraph, Table } from './model';
import { children, first, getW, type XmlElement, WORD_NS } from './xml';
import { canEditTableStructure } from './write-table';
import { isStJcTable, isStVerticalJc } from './generated/wml-simple-types';
import { enumValue } from './parse-diagnostics';
import {
	onOffAttribute,
	onOffElement,
	parseSignedTwips,
	parseTwips,
	parseUnsignedInteger,
} from './simple-types';
import type { Twips } from './units';
import { parseShadingFill, parseShadingThemeFill, parseTableBorders } from './table-borders';

const dxa = parseTwips;

function nestedPreview(node: XmlElement): NestedTablePreview {
	return {
		rows: children(node, 'tr').map((row) =>
			children(row, 'tc').map((cell) => ({
				text: Array.from(cell.getElementsByTagNameNS(WORD_NS, 't'))
					.map((t) => t.textContent ?? '')
					.join(''),
			})),
		),
	};
}

function parseCell(
	cell: XmlElement,
	parseParagraph: (p: XmlElement, id: string) => Paragraph,
	id: string,
): TableCell {
	const props = first(cell, 'tcPr');
	const result: TableCell = {
		paragraphs: children(cell, 'p').map((p, pi) => parseParagraph(p, `${id}p${pi}`)),
	};
	const gridSpan = parseUnsignedInteger(getW(first(props, 'gridSpan'), 'val'));
	if (gridSpan !== undefined) result.gridSpan = gridSpan;
	const vMerge = first(props, 'vMerge');
	if (vMerge) result.verticalMerge = getW(vMerge, 'val') === 'restart' ? 'restart' : 'continue';
	const tcW = first(props, 'tcW');
	if ((getW(tcW, 'type') ?? 'dxa') === 'dxa') {
		const width = dxa(getW(tcW, 'w'));
		if (width !== undefined) result.widthTwips = width;
	}
	const vAlign = enumValue(isStVerticalJc, getW(first(props, 'vAlign'), 'val'), 'w:tcPr/w:vAlign');
	if (vAlign) result.verticalAlign = vAlign;
	const shd = first(props, 'shd');
	const fill =
		parseShadingFill(shd) ??
		(getW(shd, 'fill') === 'auto' && !getW(shd, 'themeFill') ? 'auto' : undefined);
	if (fill) result.shadingFill = fill;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const borders = parseTableBorders(first(props, 'tcBorders'));
	if (borders) result.borders = borders;
	const margins = parseMargins(first(props, 'tcMar'));
	if (margins) result.margins = margins;
	const nested = children(cell, 'tbl').map(nestedPreview);
	if (nested.length) result.nestedTables = nested;
	return result;
}

function parseLook(element: XmlElement | undefined): TableLook | undefined {
	if (!element) return undefined;
	const flags = [
		['firstRow', 0x0020],
		['lastRow', 0x0040],
		['firstColumn', 0x0080],
		['lastColumn', 0x0100],
		['noHBand', 0x0200],
		['noVBand', 0x0400],
	] as const;
	// Word reads the transitional bitmask only when none of the named flags are present.
	const namedFlags = flags.some(([name]) => getW(element, name) !== undefined);
	const raw = getW(element, 'val');
	const packed = !namedFlags && raw && /^[0-9a-f]{4}$/i.test(raw) ? parseInt(raw, 16) : undefined;
	const look: TableLook = {};
	for (const [name, bit] of flags) {
		const value = packed === undefined ? onOffAttribute(element, name) : (packed & bit) !== 0;
		if (value !== undefined) look[name] = value;
	}
	return Object.keys(look).length ? look : undefined;
}

/** Cell margins from `w:tcMar` or `w:tblCellMar`, in twips (`start`/`end` read as left/right). */
function parseMargins(element: XmlElement | undefined): TableCellMargins | undefined {
	if (!element) return undefined;
	const parsed: TableCellMargins = {};
	for (const [side, logical] of [
		['top', 'top'],
		['bottom', 'bottom'],
		['left', 'start'],
		['right', 'end'],
	] as const) {
		const edge = first(element, logical) ?? first(element, side);
		const type = getW(edge, 'type') ?? 'dxa';
		const value =
			type === 'nil'
				? parseSignedTwips('0')
				: type === 'dxa'
					? parseSignedTwips(getW(edge, 'w'))
					: undefined;
		if (value !== undefined) parsed[side] = value;
	}
	return Object.keys(parsed).length ? parsed : undefined;
}

/** Row height (`w:trHeight`, `auto` rule meaning at least), keep-together and header flags. */
function parseRowProperties(row: XmlElement): TableRowProperties {
	const trPr = first(row, 'trPr');
	const result: TableRowProperties = {};
	const height = first(trPr, 'trHeight');
	const heightTwips = dxa(getW(height, 'val'));
	if (heightTwips) {
		result.heightTwips = heightTwips;
		result.heightRule = getW(height, 'hRule') === 'exact' ? 'exact' : 'atLeast';
	}
	if (onOffElement(first(trPr, 'cantSplit'))) result.cantSplit = true;
	if (onOffElement(first(trPr, 'tblHeader'))) result.header = true;
	return result;
}

/** Parses a `w:tbl`, including grid widths, merges, borders, shading and table style/look. */
export function parseTable(
	node: XmlElement,
	id: string,
	parseParagraph: (p: XmlElement, id: string) => Paragraph,
): Table {
	const rows = children(node, 'tr').map((row, ri) =>
		children(row, 'tc').map((cell, ci) => parseCell(cell, parseParagraph, `${id}-r${ri}c${ci}`)),
	);
	const table: Table = { type: 'table', id, rows, structureEditable: canEditTableStructure(node) };
	const gridElement = first(node, 'tblGrid');
	const grid = gridElement
		? children(gridElement, 'gridCol').flatMap((column): Twips[] => {
				const width = dxa(getW(column, 'w'));
				return width === undefined ? [] : [width];
			})
		: [];
	if (grid.length) table.grid = grid;
	const tblPr = first(node, 'tblPr');
	const tableWidth = first(tblPr, 'tblW');
	const width =
		(getW(tableWidth, 'type') ?? 'dxa') === 'dxa' ? dxa(getW(tableWidth, 'w')) : undefined;
	if (width !== undefined) table.widthTwips = width;
	const justification = enumValue(isStJcTable, getW(first(tblPr, 'jc'), 'val'), 'w:tblPr/w:jc');
	if (justification) {
		table.justification = justification;
		const rtl = onOffElement(first(tblPr, 'bidiVisual')) === true;
		table.alignment =
			justification === 'start'
				? rtl
					? 'right'
					: 'left'
				: justification === 'end'
					? rtl
						? 'left'
						: 'right'
					: justification;
	}
	const indent = parseSignedTwips(getW(first(tblPr, 'tblInd'), 'w'));
	if (indent !== undefined) table.indentTwips = indent;
	const borders = parseTableBorders(first(tblPr, 'tblBorders'));
	if (borders) table.borders = borders;
	const styleId = getW(first(tblPr, 'tblStyle'), 'val');
	if (styleId) table.style = styleId;
	const look = parseLook(first(tblPr, 'tblLook'));
	if (look) table.look = look;
	const cellMargins = parseMargins(first(tblPr, 'tblCellMar'));
	if (cellMargins) table.cellMargins = cellMargins;
	const rowProperties = children(node, 'tr').map(parseRowProperties);
	if (rowProperties.some((row) => Object.keys(row).length)) table.rowProperties = rowProperties;
	return table;
}
