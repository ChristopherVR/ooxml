import { MAX_COL, MAX_ROW } from '../address.js';
import { usedRange } from '../cells.js';
import type { Worksheet } from '../model.js';
import { AxisMetrics } from './axis-metrics.js';
import {
	DEFAULT_MAX_DIGIT_WIDTH,
	columnWidthToPixels,
	defaultColumnPixels,
	pointsToPixels,
} from './units.js';

/** Pixel geometry of a sheet's grid at a zoom level (cell area only, no headers). */
export interface GridMetrics {
	colLeft(c: number): number;
	colWidth(c: number): number;
	rowTop(r: number): number;
	rowHeight(r: number): number;
	/** The visible column under `x` (hidden columns are skipped). */
	colAt(x: number): number;
	/** The visible row under `y` (hidden rows are skipped). */
	rowAt(y: number): number;
	/** Width of columns `0..maxCol` (default: the scroll extent). */
	totalWidth(maxCol?: number): number;
	/** Height of rows `0..maxRow` (default: the scroll extent). */
	totalHeight(maxRow?: number): number;
	/** Zoom in percent (100 = actual size). */
	zoom: number;
	/** Last row / column of the suggested scroll extent (used range plus the extra margin). */
	lastRow: number;
	lastCol: number;
	isRowHidden(r: number): boolean;
	isColHidden(c: number): boolean;
}

export interface GridMetricsOptions {
	/**
	 * Zoom in percent (defaults to `sheet.view.zoom`). Values of 4 or less are read as factors
	 * (`1.5` = 150%), since Excel's smallest zoom is 10%.
	 */
	zoom?: number;
	/** Maximum digit width of the workbook's default font, in pixels. */
	mdw?: number;
	/** Rows beyond the used range included in the scroll extent (default 50). */
	extraRows?: number;
	/** Columns beyond the used range included in the scroll extent (default 20). */
	extraCols?: number;
}

/** Normalizes a zoom given as percent or factor to percent within Excel's 10%-400% range. */
export function zoomPercent(zoom: number | undefined): number {
	if (zoom === undefined || !Number.isFinite(zoom) || zoom <= 0) return 100;
	const percent = zoom <= 4 ? zoom * 100 : zoom;
	return Math.min(400, Math.max(10, percent));
}

/** Column pixel overrides (custom widths and hidden columns) at 100% zoom. */
function columnOverrides(sheet: Worksheet, mdw: number, defaultPx: number): Map<number, number> {
	const out = new Map<number, number>();
	for (const info of sheet.columns) {
		const px = info.hidden
			? 0
			: info.width !== undefined
				? columnWidthToPixels(info.width, mdw)
				: defaultPx;
		if (px === defaultPx) continue;
		const last = Math.min(info.max, MAX_COL);
		for (let c = Math.max(0, info.min); c <= last; c++) out.set(c, px);
	}
	return out;
}

/** Builds the grid geometry of a sheet. Sizes are whole pixels scaled by the zoom. */
export function createGridMetrics(sheet: Worksheet, options: GridMetricsOptions = {}): GridMetrics {
	const zoom = zoomPercent(options.zoom ?? sheet.view.zoom);
	const factor = zoom / 100;
	const mdw = options.mdw ?? DEFAULT_MAX_DIGIT_WIDTH;
	const scale = (px: number): number => (px > 0 ? Math.max(1, Math.round(px * factor)) : 0);

	const defaultColPx = defaultColumnPixels(sheet.defaultColWidth, mdw);
	const colOverrides = [...columnOverrides(sheet, mdw, defaultColPx)].map(
		([c, px]) => [c, scale(px)] as const,
	);
	const cols = new AxisMetrics(scale(defaultColPx), MAX_COL + 1, colOverrides);

	const defaultRowPx = pointsToPixels(sheet.defaultRowHeight);
	const rowOverrides: [number, number][] = [];
	for (const [r, info] of sheet.rowInfo) {
		if (info.hidden) rowOverrides.push([r, 0]);
		else if (info.height !== undefined) rowOverrides.push([r, scale(pointsToPixels(info.height))]);
	}
	const rows = new AxisMetrics(scale(defaultRowPx), MAX_ROW + 1, rowOverrides);

	const used = usedRange(sheet);
	let lastRowUsed = used?.end.row ?? 0;
	for (const r of sheet.rowInfo.keys()) lastRowUsed = Math.max(lastRowUsed, r);
	for (const merge of sheet.merges) lastRowUsed = Math.max(lastRowUsed, merge.end.row);
	let lastColUsed = used?.end.col ?? 0;
	for (const merge of sheet.merges) lastColUsed = Math.max(lastColUsed, merge.end.col);
	for (const info of sheet.columns)
		if (info.max < MAX_COL) lastColUsed = Math.max(lastColUsed, info.max);
	const lastRow = Math.min(MAX_ROW, lastRowUsed + (options.extraRows ?? 50));
	const lastCol = Math.min(MAX_COL, lastColUsed + (options.extraCols ?? 20));

	return {
		zoom,
		lastRow,
		lastCol,
		colLeft: (c) => cols.start(c),
		colWidth: (c) => cols.size(c),
		rowTop: (r) => rows.start(r),
		rowHeight: (r) => rows.size(r),
		colAt: (x) => cols.at(x),
		rowAt: (y) => rows.at(y),
		totalWidth: (maxCol = lastCol) => cols.start(Math.min(maxCol, MAX_COL) + 1),
		totalHeight: (maxRow = lastRow) => rows.start(Math.min(maxRow, MAX_ROW) + 1),
		isRowHidden: (r) => rows.size(r) === 0,
		isColHidden: (c) => cols.size(c) === 0,
	};
}
