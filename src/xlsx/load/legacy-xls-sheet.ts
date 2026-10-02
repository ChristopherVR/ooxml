// Maps one BIFF8 worksheet from the ole2 reader onto the workbook model: cells with cached values
// and formula text, styles, merges, column/row sizes, view and frozen panes, hyperlinks and comments.
import type { XlsCellValue, XlsRange, XlsSheet } from '@christophervr/ole2/legacy-excel-workbook';
import type { CellRange } from '../address.js';
import {
	cellError,
	isErrorCode,
	type Cell,
	type CellValue,
	type ColumnInfo,
	type Hyperlink,
	type RowInfo,
	type Worksheet,
} from '../model.js';
import { createWorksheet } from '../workbook.js';
import { mapColor } from './legacy-xls-styles.js';

/** Resolves a BIFF8 XF index to a model style id (0 for the default format). */
export type StyleResolver = (xf: number) => number;

export interface SheetMapResult {
	sheet: Worksheet;
	/** Formula cells whose tokens the reader could not decode (kept as cached values). */
	undecodedFormulas: number;
	externalReferences: boolean;
}

/** Points per default row in an Excel 97-2003 file when `DEFAULTROWHEIGHT` is missing. */
const XLS_DEFAULT_ROW_HEIGHT = 15;

export const toRange = (range: XlsRange): CellRange => ({
	start: { row: range.firstRow, col: range.firstCol },
	end: { row: range.lastRow, col: range.lastCol },
});

function toValue(value: XlsCellValue): CellValue {
	if (value === null || typeof value !== 'object') return value;
	return cellError(isErrorCode(value.error) ? value.error : '#N/A');
}

function putCell(sheet: Worksheet, row: number, col: number, cell: Cell): void {
	let cells = sheet.rows.get(row);
	if (!cells) {
		cells = new Map();
		sheet.rows.set(row, cells);
	}
	cells.set(col, cell);
}

/** Column width in characters including padding, from DEFCOLWIDTH (characters, no padding). */
function defaultColumnWidth(source: XlsSheet): number | undefined {
	if (source.standardWidth) return source.standardWidth;
	if (source.defaultColWidth === undefined || source.defaultColWidth === 8) return undefined;
	return Math.round((source.defaultColWidth + 5 / 7) * 100) / 100;
}

export function mapSheet(source: XlsSheet, sheetId: number, style: StyleResolver): SheetMapResult {
	const sheet = createWorksheet(source.name, sheetId);
	sheet.state = source.state;
	sheet.defaultRowHeight = source.defaultRowHeight ?? XLS_DEFAULT_ROW_HEIGHT;
	const colWidth = defaultColumnWidth(source);
	if (colWidth !== undefined) sheet.defaultColWidth = colWidth;
	let undecodedFormulas = 0;
	let externalReferences = false;

	for (const item of source.cells) {
		const styleId = style(item.xf);
		const cell: Cell = { value: toValue(item.value) };
		if (item.formula !== undefined) {
			cell.formula = item.formula;
			cell.legacyFormula = true;
			if (/\[[^\]]+\][^!]*!/.test(item.formula)) externalReferences = true;
		}
		if (item.formulaUndecoded) undecodedFormulas++;
		if (item.arrayRange) cell.arrayRange = toRange(item.arrayRange);
		if (styleId) cell.styleId = styleId;
		if (cell.value === null && !cell.formula && !cell.styleId) continue;
		putCell(sheet, item.row, item.col, cell);
	}

	sheet.merges = source.merges.map(toRange);
	sheet.columns = source.columns.map((column) => {
		const info: ColumnInfo = {
			min: column.firstCol,
			max: column.lastCol,
			width: column.width,
		};
		if (column.customWidth) info.customWidth = true;
		if (column.hidden) info.hidden = true;
		if (column.bestFit) info.bestFit = true;
		if (column.outlineLevel) info.outlineLevel = column.outlineLevel;
		if (column.collapsed) info.collapsed = true;
		const styleId = style(column.xf);
		if (styleId) info.styleId = styleId;
		return info;
	});
	for (const row of source.rows) {
		const styleId = row.xf === undefined ? 0 : style(row.xf);
		const relevant =
			row.customHeight ||
			row.hidden ||
			row.outlineLevel > 0 ||
			row.collapsed ||
			styleId > 0 ||
			row.height !== sheet.defaultRowHeight;
		if (!relevant) continue;
		const info: RowInfo = { height: row.height };
		if (row.customHeight) info.customHeight = true;
		if (row.hidden) info.hidden = true;
		if (row.outlineLevel) info.outlineLevel = row.outlineLevel;
		if (row.collapsed) info.collapsed = true;
		if (styleId) info.styleId = styleId;
		sheet.rowInfo.set(row.row, info);
	}

	const view = source.view;
	sheet.view.showGridLines = view.showGridLines;
	sheet.view.showHeaders = view.showHeaders;
	sheet.view.showZeros = view.showZeros;
	sheet.view.rightToLeft = view.rightToLeft;
	sheet.view.zoom = view.zoom ?? 100;
	if (view.freeze) {
		sheet.view.freeze = { rows: view.freeze.rows, cols: view.freeze.cols };
		sheet.view.topLeft = { row: view.freeze.topRow, col: view.freeze.leftCol };
	} else if (view.topRow || view.leftCol)
		sheet.view.topLeft = { row: view.topRow, col: view.leftCol };
	if (view.activeCell)
		sheet.view.selection = {
			active: view.activeCell,
			ranges: (view.selection ?? []).map(toRange),
		};

	sheet.hyperlinks = source.hyperlinks.map((link) => {
		const out: Hyperlink = { range: toRange(link.range) };
		if (link.target !== undefined) out.target = link.target;
		if (link.location !== undefined) out.location = link.location;
		if (link.tooltip !== undefined) out.tooltip = link.tooltip;
		if (link.display !== undefined) out.display = link.display;
		return out;
	});
	sheet.comments = source.comments.map((note) => ({
		address: { row: note.row, col: note.col },
		author: note.author,
		text: note.text,
	}));
	const tab = mapColor(source.tabColor);
	if (tab) sheet.tabColor = tab;
	if (source.protection) {
		sheet.protection = { sheet: true };
		if (source.protection.passwordHash)
			sheet.protection.passwordHash = source.protection.passwordHash;
	}
	return { sheet, undecodedFormulas, externalReferences };
}
