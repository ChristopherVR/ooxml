/**
 * A table style's own background (`a:tblBg`) is one surface under the whole
 * table: a gradient on it runs once from the top edge to the bottom edge, not
 * once per cell. Cells are painted individually here, so each cell that still
 * shows the table background gradient gets a `background-size` and
 * `background-position` that map its box onto its slice of the table.
 */
import type { PptxTableData } from 'ooxml-core/pptx';

import type { TableCellCss } from './table-style';

/** Fractions of the table's width and height covered by the cell, and its offset. */
export interface TableCellSlice {
	left: number;
	top: number;
	width: number;
	height: number;
}

/** Locate a cell inside the table as fractions of the table's size. */
export function tableCellSlice(
	tableData: PptxTableData,
	rowIndex: number,
	cellIndex: number,
): TableCellSlice | undefined {
	const rows = tableData.rows;
	const cols = tableData.columnWidths;
	if (rows.length === 0 || cols.length === 0) {
		return undefined;
	}
	const heights = rows.map((row) => row.height ?? 0);
	const useHeights = heights.every((h) => h > 0);
	const rowWeights = useHeights ? heights : rows.map(() => 1);
	const colTotal = cols.reduce((sum, w) => sum + w, 0);
	const rowTotal = rowWeights.reduce((sum, w) => sum + w, 0);
	if (colTotal <= 0 || rowTotal <= 0) {
		return undefined;
	}
	const before = (weights: number[], index: number): number =>
		weights.slice(0, index).reduce((sum, w) => sum + w, 0);
	return {
		left: before(cols, cellIndex) / colTotal,
		top: before(rowWeights, rowIndex) / rowTotal,
		width: (cols[cellIndex] ?? 0) / colTotal,
		height: (rowWeights[rowIndex] ?? 0) / rowTotal,
	};
}

/**
 * Make the table-background gradient `tableBackgroundCss` span the whole table
 * instead of restarting in each cell. A no-op unless the cell's final
 * `background` still ends with that gradient (a translucent band tint is
 * layered on top of it; an opaque fill replaced it).
 */
export function spanTableBackground(
	css: TableCellCss,
	tableBackgroundCss: string | undefined,
	tableData: PptxTableData,
	rowIndex: number,
	cellIndex: number,
): void {
	const background = css.background;
	if (
		!tableBackgroundCss ||
		typeof background !== 'string' ||
		!background.endsWith(tableBackgroundCss) ||
		!tableBackgroundCss.includes('gradient(')
	) {
		return;
	}
	const slice = tableCellSlice(tableData, rowIndex, cellIndex);
	if (!slice || slice.width <= 0 || slice.height <= 0) {
		return;
	}
	const pct = (value: number): string => `${Math.round(value * 1e4) / 100}%`;
	const size = `${pct(1 / slice.width)} ${pct(1 / slice.height)}`;
	const x = slice.width >= 1 ? 0 : slice.left / (1 - slice.width);
	const y = slice.height >= 1 ? 0 : slice.top / (1 - slice.height);
	const position = `${pct(x)} ${pct(y)}`;
	const layers = background.slice(0, background.length - tableBackgroundCss.length);
	// Layers above the table gradient (a translucent tint) keep their own box.
	const upper = layers.trim().length > 0;
	css.backgroundSize = upper ? `auto, ${size}` : size;
	css.backgroundPosition = upper ? `0% 0%, ${position}` : position;
	css.backgroundRepeat = 'no-repeat';
}
