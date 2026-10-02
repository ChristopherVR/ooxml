// Pure pane geometry: where the frozen and scrolling quadrants sit, how a cell maps to view
// pixels and back (hit testing), and header sizes. View coordinates have (0,0) at the top-left of
// the grid element, headers included. Cell coordinates come from the core's GridMetrics.
import type { CellRange, FreezePane, GridMetrics } from '@christophervr/xlsx-core';

export type Quadrant = 'corner' | 'top' | 'left' | 'main';
export const QUADRANTS: readonly Quadrant[] = ['main', 'top', 'left', 'corner'];

export interface Box {
	x: number;
	y: number;
	w: number;
	h: number;
}

export interface GeometryInput {
	metrics: GridMetrics;
	freeze: FreezePane | undefined;
	headerW: number;
	headerH: number;
	/** Size of the whole grid element (headers included). */
	width: number;
	height: number;
	scrollLeft: number;
	scrollTop: number;
}

export type HitArea = 'cell' | 'col-header' | 'row-header' | 'corner' | 'outside';
export interface Hit {
	area: HitArea;
	row: number;
	col: number;
}

export class GridGeometry {
	readonly metrics: GridMetrics;
	readonly fRows: number;
	readonly fCols: number;
	readonly frozenW: number;
	readonly frozenH: number;
	readonly headerW: number;
	readonly headerH: number;
	readonly width: number;
	readonly height: number;
	readonly scrollLeft: number;
	readonly scrollTop: number;

	constructor(input: GeometryInput) {
		this.metrics = input.metrics;
		this.fRows = Math.max(0, input.freeze?.rows ?? 0);
		this.fCols = Math.max(0, input.freeze?.cols ?? 0);
		this.frozenW = this.fCols > 0 ? input.metrics.colLeft(this.fCols) : 0;
		this.frozenH = this.fRows > 0 ? input.metrics.rowTop(this.fRows) : 0;
		this.headerW = input.headerW;
		this.headerH = input.headerH;
		this.width = input.width;
		this.height = input.height;
		this.scrollLeft = Math.max(0, input.scrollLeft);
		this.scrollTop = Math.max(0, input.scrollTop);
	}

	/** Width / height of the cell area (no headers). */
	get cellsW(): number {
		return Math.max(0, this.width - this.headerW);
	}
	get cellsH(): number {
		return Math.max(0, this.height - this.headerH);
	}

	/** The box of a quadrant in view coordinates, clipped to the element. */
	box(q: Quadrant): Box {
		const fw = Math.min(this.frozenW, this.cellsW);
		const fh = Math.min(this.frozenH, this.cellsH);
		const left = q === 'corner' || q === 'left';
		const top = q === 'corner' || q === 'top';
		return {
			x: this.headerW + (left ? 0 : fw),
			y: this.headerH + (top ? 0 : fh),
			w: left ? fw : Math.max(0, this.cellsW - fw),
			h: top ? fh : Math.max(0, this.cellsH - fh),
		};
	}

	/** Plane offset of a quadrant: a cell at (colLeft, rowTop) shows at box + (left - ox, top - oy). */
	origin(q: Quadrant): { ox: number; oy: number } {
		const scrollX = q === 'main' || q === 'top';
		const scrollY = q === 'main' || q === 'left';
		return {
			ox: scrollX ? this.frozenW + this.scrollLeft : 0,
			oy: scrollY ? this.frozenH + this.scrollTop : 0,
		};
	}

	/** The quadrant a cell is painted in. */
	quadrantOf(row: number, col: number): Quadrant {
		const top = row < this.fRows;
		const left = col < this.fCols;
		return top ? (left ? 'corner' : 'top') : left ? 'left' : 'main';
	}

	/** View x of a column's left edge. */
	colX(col: number): number {
		const left = this.metrics.colLeft(col);
		return this.headerW + (col < this.fCols ? left : left - this.scrollLeft);
	}

	rowY(row: number): number {
		const top = this.metrics.rowTop(row);
		return this.headerH + (row < this.fRows ? top : top - this.scrollTop);
	}

	/** View rectangle of a range (spanning panes is approximated from its top-left cell's pane). */
	rangeBox(range: CellRange): Box {
		const x = this.colX(range.start.col);
		const y = this.rowY(range.start.row);
		const w = this.metrics.colLeft(range.end.col + 1) - this.metrics.colLeft(range.start.col);
		const h = this.metrics.rowTop(range.end.row + 1) - this.metrics.rowTop(range.start.row);
		return { x, y, w, h };
	}

	/** Column under a view x (clamped into the cell area). */
	colAtX(x: number): number {
		const cx = Math.max(0, x - this.headerW);
		if (cx < this.frozenW) return this.metrics.colAt(cx);
		return this.metrics.colAt(cx + this.scrollLeft);
	}

	rowAtY(y: number): number {
		const cy = Math.max(0, y - this.headerH);
		if (cy < this.frozenH) return this.metrics.rowAt(cy);
		return this.metrics.rowAt(cy + this.scrollTop);
	}

	hit(x: number, y: number): Hit {
		if (x < 0 || y < 0 || x > this.width || y > this.height)
			return { area: 'outside', row: -1, col: -1 };
		const inHeaderX = x < this.headerW;
		const inHeaderY = y < this.headerH;
		if (inHeaderX && inHeaderY) return { area: 'corner', row: -1, col: -1 };
		if (inHeaderY) return { area: 'col-header', row: -1, col: this.colAtX(x) };
		if (inHeaderX) return { area: 'row-header', row: this.rowAtY(y), col: -1 };
		return { area: 'cell', row: this.rowAtY(y), col: this.colAtX(x) };
	}

	/**
	 * A column border near `x` in the column header (for resizing): the column whose right edge
	 * is within `slop` pixels, or undefined.
	 */
	colBorderAt(x: number, slop = 4): number | undefined {
		const col = this.colAtX(x);
		const right = this.colX(col) + this.metrics.colWidth(col);
		if (Math.abs(x - right) <= slop) return col;
		const left = this.colX(col);
		if (Math.abs(x - left) <= slop && col > 0) {
			let prev = col - 1;
			while (prev > 0 && this.metrics.colWidth(prev) === 0) prev--;
			return prev;
		}
		return undefined;
	}

	rowBorderAt(y: number, slop = 3): number | undefined {
		const row = this.rowAtY(y);
		const bottom = this.rowY(row) + this.metrics.rowHeight(row);
		if (Math.abs(y - bottom) <= slop) return row;
		const top = this.rowY(row);
		if (Math.abs(y - top) <= slop && row > 0) {
			let prev = row - 1;
			while (prev > 0 && this.metrics.rowHeight(prev) === 0) prev--;
			return prev;
		}
		return undefined;
	}
}

/** Row header width for the largest row number in view (Excel grows it with the digits). */
export function rowHeaderWidth(lastRow: number, zoom: number): number {
	const digits = String(lastRow + 1).length;
	return Math.round((Math.max(2, digits) * 7 + 12) * (zoom / 100));
}

export const columnHeaderHeight = (zoom: number): number => Math.round(20 * (zoom / 100));

/** The scroll offsets that bring a cell fully into the scrolling pane, or undefined if it is visible. */
export function scrollToReveal(
	g: GridGeometry,
	row: number,
	col: number,
): { left: number; top: number } {
	let { scrollLeft: left, scrollTop: top } = g;
	const m = g.metrics;
	if (col >= g.fCols) {
		const x = m.colLeft(col) - g.frozenW;
		const w = m.colWidth(col);
		const viewW = Math.max(0, g.cellsW - g.frozenW);
		if (x < left) left = x;
		else if (x + w > left + viewW) left = Math.min(x, x + w - viewW);
	}
	if (row >= g.fRows) {
		const y = m.rowTop(row) - g.frozenH;
		const h = m.rowHeight(row);
		const viewH = Math.max(0, g.cellsH - g.frozenH);
		if (y < top) top = y;
		else if (y + h > top + viewH) top = Math.min(y, y + h - viewH);
	}
	return { left: Math.max(0, left), top: Math.max(0, top) };
}
