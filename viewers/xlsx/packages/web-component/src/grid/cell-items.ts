// Builds the paint list of one quadrant: which cells get a node, at what size (merges), with what
// final text (#### for numbers that do not fit, repeated text for `fill`) and how far text spills
// into empty neighbours (core `overflowExtent`). Pure apart from the injected view and measurer.
import {
	overflowExtent,
	type CellRange,
	type CellView,
	type GridMetrics,
	type Worksheet,
} from '@christophervr/xlsx-core';
import { cssFont, hashes, hasBorders, repeatToFill } from './cell-paint.js';

export interface CellItem {
	key: string;
	row: number;
	col: number;
	/** Plane pixels (merge-sized). */
	x: number;
	y: number;
	w: number;
	h: number;
	view: CellView;
	text: string;
	/** Text box relative to the cell's left edge (wider than the cell when text overflows). */
	textLeft: number;
	textWidth: number;
	overflowing: boolean;
	/** Font scale for shrink-to-fit (1 = none). */
	fontScale: number;
	merged: boolean;
	errorMark: boolean;
}

export interface BuildInput {
	sheet: Worksheet;
	metrics: GridMetrics;
	zoom: number;
	rows: readonly number[];
	cols: readonly number[];
	/** Cached core cell view; null when the cell paints nothing. */
	view(row: number, col: number): CellView | null;
	measure(text: string, font: string): number;
	/** Formula cells with an error value get Excel's green error-check corner. */
	hasFormula(row: number, col: number): boolean;
	/**
	 * The number's text fitted to `width` CSS px as Excel does (General drops decimals or goes
	 * scientific, anything else becomes `###`), from the core's `formatValue` width option.
	 */
	fitNumber?(row: number, col: number, width: number, font: string): string | undefined;
}

const PADDING = 2;

/** True when a view paints nothing at all (no node needed). */
export function isBlankView(view: CellView): boolean {
	return (
		view.text === '' &&
		!view.fill &&
		!hasBorders(view.borders) &&
		!view.dataBar &&
		!view.icon &&
		!view.hasComment
	);
}

function textWidthOf(input: BuildInput, view: CellView, text: string): number {
	if (view.rich?.length)
		return view.rich.reduce(
			(sum, run) => sum + input.measure(run.text, cssFont(run.font, input.zoom)),
			0,
		);
	const font = cssFont(view.font, input.zoom);
	let widest = 0;
	for (const line of text.split('\n')) widest = Math.max(widest, input.measure(line, font));
	return widest;
}

function makeItem(
	input: BuildInput,
	row: number,
	col: number,
	view: CellView,
	merge: CellRange | undefined,
): CellItem {
	const { metrics } = input;
	const x = metrics.colLeft(col);
	const y = metrics.rowTop(row);
	const w = merge ? metrics.colLeft(merge.end.col + 1) - x : metrics.colWidth(col);
	const h = merge ? metrics.rowTop(merge.end.row + 1) - y : metrics.rowHeight(row);
	const item: CellItem = {
		key: `${row}:${col}`,
		row,
		col,
		x,
		y,
		w,
		h,
		view,
		text: view.text,
		textLeft: 0,
		textWidth: w,
		overflowing: false,
		fontScale: 1,
		merged: Boolean(merge),
		errorMark: view.isError && input.hasFormula(row, col),
	};
	if (!view.text || view.wrap || view.verticalText || view.rotation) return item;
	const inner = w - 2 * PADDING - view.indentPx * (input.zoom / 100);
	const measured = textWidthOf(input, view, view.text);
	if (view.hAlign === 'fill') {
		item.text = repeatToFill(view.text, measured, inner);
		return item;
	}
	if (view.shrink) {
		if (measured > inner && measured > 0) item.fontScale = Math.max(0.1, inner / measured);
		return item;
	}
	if (view.isNumber && measured > inner && !view.rich) {
		const font = cssFont(view.font, input.zoom);
		item.text = input.fitNumber?.(row, col, inner, font) ?? hashes(inner, input.measure('#', font));
		return item;
	}
	if (view.overflow && !merge && measured > inner) {
		const extent = overflowExtent(
			input.sheet,
			metrics,
			row,
			col,
			measured + 2 * PADDING,
			view.hAlign,
		);
		item.textLeft = extent.left - x;
		item.textWidth = extent.right - extent.left;
		item.overflowing = extent.startCol !== col || extent.endCol !== col;
	}
	return item;
}

/** Every merge intersecting the rows x cols block. */
function mergesIn(sheet: Worksheet, rows: readonly number[], cols: readonly number[]): CellRange[] {
	if (!sheet.merges.length || !rows.length || !cols.length) return [];
	const r0 = rows[0] ?? 0;
	const r1 = rows[rows.length - 1] ?? r0;
	const c0 = cols[0] ?? 0;
	const c1 = cols[cols.length - 1] ?? c0;
	return sheet.merges.filter(
		(m) => m.start.row <= r1 && m.end.row >= r0 && m.start.col <= c1 && m.end.col >= c0,
	);
}

export function buildItems(input: BuildInput): CellItem[] {
	const { sheet, rows, cols } = input;
	const out: CellItem[] = [];
	if (!rows.length || !cols.length) return out;
	const merges = mergesIn(sheet, rows, cols);
	const covered = (row: number, col: number): boolean =>
		merges.some(
			(m) => row >= m.start.row && row <= m.end.row && col >= m.start.col && col <= m.end.col,
		);
	for (const merge of merges) {
		const view = input.view(merge.start.row, merge.start.col);
		if (view) out.push(makeItem(input, merge.start.row, merge.start.col, view, merge));
		else
			out.push(
				makeItem(input, merge.start.row, merge.start.col, emptyMergeView(input, merge), merge),
			);
	}
	const firstCol = cols[0] ?? 0;
	const lastCol = cols[cols.length - 1] ?? firstCol;
	for (const row of rows) {
		for (const col of cols) {
			if (merges.length && covered(row, col)) continue;
			const view = input.view(row, col);
			if (view) out.push(makeItem(input, row, col, view, undefined));
		}
		// Text stored outside the painted columns may still spill into them.
		const stored = sheet.rows.get(row);
		if (!stored) continue;
		for (const [col, cell] of stored) {
			if (col >= firstCol && col <= lastCol) continue;
			if (typeof cell.value !== 'string' || cell.value === '') continue;
			if (Math.abs(col < firstCol ? firstCol - col : col - lastCol) > 64) continue;
			const view = input.view(row, col);
			if (!view?.overflow) continue;
			const item = makeItem(input, row, col, view, undefined);
			const left = item.x + item.textLeft;
			const right = left + item.textWidth;
			const viewLeft = input.metrics.colLeft(firstCol);
			const viewRight = input.metrics.colLeft(lastCol + 1);
			if (item.overflowing && right > viewLeft && left < viewRight) out.push(item);
		}
	}
	return out;
}

/** A merged area without a stored anchor still hides the gridlines inside it. */
function emptyMergeView(input: BuildInput, merge: CellRange): CellView {
	const view = input.view(merge.start.row, merge.start.col);
	if (view) return view;
	return {
		text: '',
		font: {
			family: 'Calibri',
			sizePx: 14.67,
			bold: false,
			italic: false,
			strike: false,
			color: '#000000',
		},
		borders: {},
		hAlign: 'left',
		vAlign: 'bottom',
		wrap: false,
		shrink: false,
		indentPx: 0,
		rotation: 0,
		isNumber: false,
		overflow: false,
		hasComment: false,
		hasHyperlink: false,
		validationList: false,
		isError: false,
	};
}
