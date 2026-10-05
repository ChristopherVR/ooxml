import { MAX_COL, MAX_ROW } from '../address.js';
import type { FreezePane } from '../model.js';
import type { GridMetrics } from './metrics.js';

/**
 * The scrollable pane's scroll offsets and the size of the whole cell area (frozen panes included),
 * in pixels, excluding row and column headers.
 */
export interface Viewport {
	scrollLeft: number;
	scrollTop: number;
	width: number;
	height: number;
}

export interface VisibleCells {
	/** Scrolling rows in view, ascending, hidden rows left out. */
	rows: number[];
	cols: number[];
	/** Frozen rows / columns (always painted), hidden ones left out. */
	frozenRows: number[];
	frozenCols: number[];
}

/** Frozen extent in pixels along one axis. */
const frozenExtent = (count: number, start: (i: number) => number): number =>
	count > 0 ? start(count) : 0;

function collect(
	from: number,
	to: number,
	limit: number,
	at: (offset: number) => number,
	start: (i: number) => number,
	size: (i: number) => number,
	minIndex: number,
): number[] {
	const out: number[] = [];
	if (to <= from) return out;
	let i = Math.max(minIndex, at(from));
	while (i <= limit && start(i) < to) {
		if (size(i) > 0) out.push(i);
		i++;
	}
	return out;
}

function frozenList(count: number, size: (i: number) => number, limit: number): number[] {
	const out: number[] = [];
	for (let i = 0; i < Math.min(count, limit + 1); i++) if (size(i) > 0) out.push(i);
	return out;
}

/**
 * The rows and columns to paint. With frozen panes, `scrollTop`/`scrollLeft` scroll only the area
 * below and right of the frozen rows and columns, exactly like Excel; the first scrolling row is
 * the one at `rowTop(freeze.rows) + scrollTop`.
 */
export function visibleCells(
	metrics: GridMetrics,
	viewport: Viewport,
	freeze?: FreezePane,
): VisibleCells {
	const fRows = Math.max(0, freeze?.rows ?? 0);
	const fCols = Math.max(0, freeze?.cols ?? 0);
	const frozenH = frozenExtent(fRows, metrics.rowTop);
	const frozenW = frozenExtent(fCols, metrics.colLeft);
	const top = frozenH + Math.max(0, viewport.scrollTop);
	const left = frozenW + Math.max(0, viewport.scrollLeft);
	const rows = collect(
		top,
		top + Math.max(0, viewport.height - frozenH),
		MAX_ROW,
		metrics.rowAt,
		metrics.rowTop,
		metrics.rowHeight,
		fRows,
	);
	const cols = collect(
		left,
		left + Math.max(0, viewport.width - frozenW),
		MAX_COL,
		metrics.colAt,
		metrics.colLeft,
		metrics.colWidth,
		fCols,
	);
	return {
		rows,
		cols,
		frozenRows: frozenList(fRows, metrics.rowHeight, MAX_ROW),
		frozenCols: frozenList(fCols, metrics.colWidth, MAX_COL),
	};
}
