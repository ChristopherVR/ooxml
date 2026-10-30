// Split window maths (`pane state="split"`): sizes are stored in twentieths of a point, panes are
// laid out in pixels, and every pane scrolls on its own. DOM-free.
import { MAX_COL, MAX_ROW, type CellAddress } from '../address';
import type { SplitPane, SplitPaneName } from '../model';
import type { GridMetrics } from './metrics';
import { collectVisible, type VisibleCells } from './viewport';

/** One CSS pixel (96 per inch) is 0.75 pt, which is 15 twentieths of a point. */
const TWIPS_PER_PX = 15;
/** Narrowest pane a split bar can leave (pixels, before zoom). */
export const MIN_SPLIT_PANE = 24;

/** The saved active pane when it exists, otherwise the last scrollable pane. */
export function splitActivePane(split: Pick<SplitPane, 'x' | 'y' | 'activePane'>): SplitPaneName {
	const valid: SplitPaneName[] =
		split.x > 0 && split.y > 0
			? ['topLeft', 'topRight', 'bottomLeft', 'bottomRight']
			: split.y > 0
				? ['topLeft', 'bottomLeft']
				: ['topLeft', 'topRight'];
	return split.activePane && valid.includes(split.activePane)
		? split.activePane
		: valid[valid.length - 1]!;
}

/** Pixels (at `zoom` percent) of a split size in twentieths of a point. */
export const splitToPx = (twips: number, zoom: number): number =>
	twips > 0 ? Math.round(((twips / TWIPS_PER_PX) * zoom) / 100) : 0;

/** Twentieths of a point for a pixel distance at `zoom` percent. */
export const pxToSplit = (px: number, zoom: number): number =>
	px > 0 ? Math.round((px * TWIPS_PER_PX * 100) / zoom) : 0;

/** Scroll offsets, in plane pixels, of the panes that do not follow the native scroller. */
export interface SplitScroll {
	/** Left edge of the left panes. */
	left: number;
	/** Top edge of the top panes. */
	top: number;
}

/** The split Excel makes for View > Split with `active` selected, or undefined when it cannot. */
export function splitAtCell(
	metrics: GridMetrics,
	active: CellAddress,
	scroll: { left: number; top: number },
	view: { width: number; height: number },
): SplitPane | undefined {
	if (view.width <= 0 || view.height <= 0) return undefined;
	const ax = metrics.colLeft(active.col) - scroll.left;
	const ay = metrics.rowTop(active.row) - scroll.top;
	let x = ax > 0 && ax < view.width ? ax : 0;
	let y = ay > 0 && ay < view.height ? ay : 0;
	// Nothing above or left of the cell (A1): Excel puts the cross in the middle of the window.
	if (!x && !y) {
		x = Math.floor(view.width / 2);
		y = Math.floor(view.height / 2);
	}
	const first = {
		row: metrics.rowAt(scroll.top),
		col: metrics.colAt(scroll.left),
	};
	const split: SplitPane = {
		x: pxToSplit(x, metrics.zoom),
		y: pxToSplit(y, metrics.zoom),
		topLeft: first,
		activePane: x && y ? 'bottomRight' : y ? 'bottomLeft' : 'topRight',
	};
	return split;
}

/** Clamps a bar position (pixels) so both panes stay at least `MIN_SPLIT_PANE` wide or high. */
export function clampSplitBar(px: number, extent: number, zoom: number): number {
	const min = Math.round((MIN_SPLIT_PANE * zoom) / 100);
	return Math.max(min, Math.min(extent - min, Math.round(px)));
}

export interface SplitLayout {
	/** Left and top pane sizes in pixels (0 when not split along that axis). */
	x: number;
	y: number;
	/** Scroll of the left / top panes. */
	left: number;
	top: number;
	/** Scroll of the right / bottom panes (the native scroller). */
	mainLeft: number;
	mainTop: number;
	/** Cell area size, headers excluded. */
	width: number;
	height: number;
}

/**
 * The rows and columns each pane paints, in the shape the frozen-pane painter uses: `frozenRows` and
 * `frozenCols` are the top and left panes, `rows` and `cols` the bottom and right ones.
 */
export function splitVisible(metrics: GridMetrics, layout: SplitLayout): VisibleCells {
	const along = (
		from: number,
		size: number,
		at: (offset: number) => number,
		start: (i: number) => number,
		extent: (i: number) => number,
		limit: number,
	) => collectVisible(from, from + Math.max(0, size), limit, at, start, extent, 0);
	const rest = (total: number, taken: number) => Math.max(0, total - taken);
	return {
		frozenRows: along(
			layout.top,
			layout.y,
			metrics.rowAt,
			metrics.rowTop,
			metrics.rowHeight,
			MAX_ROW,
		),
		frozenCols: along(
			layout.left,
			layout.x,
			metrics.colAt,
			metrics.colLeft,
			metrics.colWidth,
			MAX_COL,
		),
		rows: along(
			layout.mainTop,
			rest(layout.height, layout.y),
			metrics.rowAt,
			metrics.rowTop,
			metrics.rowHeight,
			MAX_ROW,
		),
		cols: along(
			layout.mainLeft,
			rest(layout.width, layout.x),
			metrics.colAt,
			metrics.colLeft,
			metrics.colWidth,
			MAX_COL,
		),
	};
}

/**
 * The native scroll offsets that make `row`/`col` visible in some pane: unchanged when a pane
 * already shows it, else the right / bottom pane scrolls to it.
 */
export function revealInSplit(
	metrics: GridMetrics,
	layout: SplitLayout,
	row: number,
	col: number,
): { left: number; top: number } {
	const axis = (
		start: number,
		size: number,
		paneStart: number,
		paneSize: number,
		mainStart: number,
		mainSize: number,
	): number => {
		const inPane = paneSize > 0 && start >= paneStart && start + size <= paneStart + paneSize;
		if (inPane) return mainStart;
		if (start < mainStart) return start;
		if (start + size > mainStart + mainSize)
			return Math.max(0, Math.min(start, start + size - mainSize));
		return mainStart;
	};
	return {
		left: axis(
			metrics.colLeft(col),
			metrics.colWidth(col),
			layout.left,
			layout.x,
			layout.mainLeft,
			Math.max(0, layout.width - layout.x),
		),
		top: axis(
			metrics.rowTop(row),
			metrics.rowHeight(row),
			layout.top,
			layout.y,
			layout.mainTop,
			Math.max(0, layout.height - layout.y),
		),
	};
}
