import { parsePercentageShading } from './percentage-shading';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { buildXml } from '../xml/index';
import type { NestedTablePreview, TableCell, TableLook, TableRowProperties } from './table-model';
import type { Paragraph, Revision, Table } from './model';
import { children, first, getW, type XmlElement, WORD_NS } from './xml';
import { canEditTableStructure, canRestructureTable } from './table-structure-check';
import { tableGrid } from './table-grid';
import { rowRevision } from './parse-revisions';
import type { ContentControl } from './content-control-model';
import { flatCellParagraphs, flatCells, flatRows } from './parse-content-control';
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
import { parseTableFloating } from './parse-table-floating';
import { parseCellMargins } from './table-margins';

const dxa = parseTwips;

/**
 * A percentage width (`ST_MeasurementOrPercent`) as a fraction: fiftieths of a percent in the
 * transitional schema (`2500` is 50%), or a `%` value (`50%`) in the strict one.
 */
function percentWidth(width: XmlElement | undefined): number | undefined {
	if (getW(width, 'type') !== 'pct') return undefined;
	const raw = getW(width, 'w')?.trim();
	if (!raw) return undefined;
	const value = raw.endsWith('%') ? Number(raw.slice(0, -1)) / 100 : Number(raw) / 5000;
	return Number.isFinite(value) && value > 0 ? value : undefined;
}

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
	control?: ContentControl,
): TableCell {
	const props = first(cell, 'tcPr');
	const result: TableCell = {
		paragraphs: flatCellParagraphs(cell).map(({ element, control: paragraphControl }, pi) => {
			const paragraph = parseParagraph(element, `${id}p${pi}`);
			if (paragraphControl) paragraph.contentControl = paragraphControl;
			return paragraph;
		}),
	};
	if (control) result.contentControl = control;
	const gridSpan = parseUnsignedInteger(getW(first(props, 'gridSpan'), 'val'));
	if (gridSpan !== undefined) result.gridSpan = gridSpan;
	const vMerge = first(props, 'vMerge');
	if (vMerge) result.verticalMerge = getW(vMerge, 'val') === 'restart' ? 'restart' : 'continue';
	const tcW = first(props, 'tcW');
	if ((getW(tcW, 'type') ?? 'dxa') === 'dxa') {
		const width = dxa(getW(tcW, 'w'));
		if (width !== undefined) result.widthTwips = width;
	}
	const percent = percentWidth(tcW);
	if (percent !== undefined) result.widthPercent = percent;
	const vAlign = enumValue(isStVerticalJc, getW(first(props, 'vAlign'), 'val'), 'w:tcPr/w:vAlign');
	if (vAlign) result.verticalAlign = vAlign;
	const hideMark = onOffElement(first(props, 'hideMark'));
	if (hideMark !== undefined) result.hideMark = hideMark;
	const shd = first(props, 'shd');
	const fill =
		parseShadingFill(shd) ??
		(getW(shd, 'fill') === 'auto' && !getW(shd, 'themeFill') ? 'auto' : undefined);
	if (fill) result.shadingFill = fill;
	const pattern = parsePercentageShading(shd);
	if (pattern) result.percentageShading = pattern;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const borders = parseTableBorders(first(props, 'tcBorders'));
	if (borders) result.borders = borders;
	const margins = parseCellMargins(first(props, 'tcMar'));
	if (margins) result.margins = margins;
	const cellFormatRevision = parsePropertyChange(props, 'tcPrChange', 'tcPr', 'cellChange');
	if (cellFormatRevision) result.formatRevision = cellFormatRevision;
	const nested = children(cell, 'tbl').map(nestedPreview);
	if (nested.length) result.nestedTables = nested;
	if (nested.length) {
		let paragraphIndex = 0,
			tableIndex = 0;
		result.sourceBlocks = Array.from(cell.childNodes).flatMap(
			(child): import('./model').Block[] => {
				if (child.nodeType !== 1) return [];
				const element = child as XmlElement;
				if (element.namespaceURI !== WORD_NS) return [];
				if (element.localName === 'p') return [result.paragraphs[paragraphIndex++]!];
				if (element.localName === 'tbl') {
					const inner = parseTable(element, `${id}-nested${tableIndex++}`, parseParagraph);
					// A table in a cell is always edited through the merge-aware structure editors.
					markRestructurable(inner, element);
					return [inner];
				}
				return [];
			},
		);
	}
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

function parsePropertyChange(
	parent: XmlElement | undefined,
	changeTag: string,
	priorTag: string,
	kind: 'tableChange' | 'rowChange' | 'cellChange',
): Revision | undefined {
	const change = first(parent, changeTag);
	if (!change) return undefined;
	const author = getW(change, 'author') ?? '';
	const date = getW(change, 'date');
	const dateUtc = change.getAttributeNS(
		'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
		'dateUtc',
	);
	const prior = first(change, priorTag);
	return {
		id: getW(change, 'id') ?? '',
		kind,
		author,
		...(date ? { date } : {}),
		...(dateUtc ? { dateUtc } : {}),
		...(prior ? { previousPropertiesXml: buildXml(prior) } : {}),
	};
}

/** Row height (`w:trHeight`, `auto` rule meaning at least), keep-together and header flags. */
function parseRowProperties(row: XmlElement): TableRowProperties {
	const trPr = first(row, 'trPr');
	const result: TableRowProperties = {};
	const height = first(trPr, 'trHeight');
	const heightTwips = dxa(getW(height, 'val'));
	if (heightTwips !== undefined) {
		result.heightTwips = heightTwips;
		result.heightRule = getW(height, 'hRule') === 'exact' ? 'exact' : 'atLeast';
	}
	if (onOffElement(first(trPr, 'cantSplit'))) result.cantSplit = true;
	if (onOffElement(first(trPr, 'tblHeader'))) result.header = true;
	const revision = rowRevision(trPr);
	if (revision) result.revision = revision;
	const formatRevision = parsePropertyChange(trPr, 'trPrChange', 'trPr', 'rowChange');
	if (formatRevision) result.formatRevision = formatRevision;
	return result;
}

/** Flags a table whose cells tile its grid and whose XML the structure editors can keep. */
function markRestructurable(table: Table, node: XmlElement): void {
	if (canRestructureTable(node) && tableGrid(table)) table.restructurable = true;
}

/** Parses a `w:tbl`, including grid widths, merges, borders, shading and table style/look. */
export function parseTable(
	node: XmlElement,
	id: string,
	parseParagraph: (p: XmlElement, id: string) => Paragraph,
): Table {
	const rowEntries = flatRows(node);
	const rows = rowEntries.map((row, ri) =>
		flatCells(row.element).map((cell, ci) =>
			parseCell(cell.element, parseParagraph, `${id}-r${ri}c${ci}`, cell.control),
		),
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
	const percent = percentWidth(tableWidth);
	if (percent !== undefined) table.widthPercent = percent;
	table.layout = getW(first(tblPr, 'tblLayout'), 'type') === 'fixed' ? 'fixed' : 'autofit';
	const floating = parseTableFloating(first(tblPr, 'tblpPr'));
	if (floating) table.floating = floating;
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
	const cellMargins = parseCellMargins(first(tblPr, 'tblCellMar'));
	if (cellMargins) table.cellMargins = cellMargins;
	const tableFormatRevision = parsePropertyChange(tblPr, 'tblPrChange', 'tblPr', 'tableChange');
	if (tableFormatRevision) table.formatRevision = tableFormatRevision;
	const rowProperties = rowEntries.map(({ element, control }) => {
		const properties = parseRowProperties(element);
		if (control) properties.contentControl = control;
		return properties;
	});
	if (rowProperties.some((row) => Object.keys(row).length)) table.rowProperties = rowProperties;
	if (!table.structureEditable) markRestructurable(table, node);
	return table;
}
