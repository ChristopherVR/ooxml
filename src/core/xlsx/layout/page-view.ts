import { MAX_COL, MAX_ROW, type CellRange } from '../address';
import type { Workbook } from '../model';
import type { HeaderFooterSection } from './header-footer';
import { createGridMetrics } from './metrics';
import { paginateWorksheet, type PrintLayout } from './print-layout';

/** CSS pixels between neighbouring pages in Page Layout view. */
export const PAGE_VIEW_GAP = 24;

export interface PageViewRect {
	x: number;
	y: number;
	width: number;
	height: number;
}

export interface PageViewPage {
	number: number;
	/** The page's place in the view's grid of pages. */
	slotRow: number;
	slotCol: number;
	/** The whole sheet of paper. */
	paper: PageViewRect;
	/** The printable area inside the margins. */
	content: PageViewRect;
	header: PageViewRect;
	footer: PageViewRect;
	range: CellRange | undefined;
	rows: number[];
	columns: number[];
	/** Offsets of each row and column from the content origin, at 100% (before `scale`). */
	rowTops: number[];
	colLefts: number[];
	/** True when the page has no data (Excel shows "Click to add data"). */
	blank: boolean;
}

export interface PageViewLayout {
	pages: PageViewPage[];
	scale: number;
	width: number;
	height: number;
	plan: PrintLayout;
}

type Sizes = { rowHeight(r: number): number; colWidth(c: number): number };

const offsets = (indices: number[], size: (i: number) => number): number[] => {
	let at = 0;
	return indices.map((index) => {
		const here = at;
		at += size(index);
		return here;
	});
};

/** Appends the indices after `list` that still fit in `capacity` pixels (hidden ones skipped). */
function fillTo(
	list: number[],
	size: (i: number) => number,
	hidden: (i: number) => boolean,
	capacity: number,
	limit: number,
): number[] {
	const out = [...list];
	let total = out.reduce((sum, i) => sum + size(i), 0);
	for (let next = (out.at(-1) ?? -1) + 1; next <= limit; next++) {
		if (hidden(next)) continue;
		const extent = size(next);
		if (total + extent > capacity + 0.001) break;
		out.push(next);
		total += extent;
	}
	return out;
}

/** Page Layout view geometry: every page placed on a canvas whose origin is (0, 0). */
export function pageViewLayout(
	workbook: Workbook,
	sheetIndex: number,
	gap = PAGE_VIEW_GAP,
): PageViewLayout {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) throw new RangeError('Invalid worksheet.');
	const plan = paginateWorksheet(workbook, sheetIndex);
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const { paperWidth, paperHeight, margins } = plan;
	const sorted = (values: number[]) => [...new Set(values)].sort((a, b) => a - b);
	const rowSlots = sorted(plan.pages.map((page) => page.range.start.row));
	const colSlots = sorted(plan.pages.map((page) => page.range.start.col));
	const entries: {
		number: number;
		slotRow: number;
		slotCol: number;
		range: CellRange | undefined;
		rows: number[];
		columns: number[];
	}[] = plan.pages.map((page) => ({
		number: page.number,
		slotRow: rowSlots.indexOf(page.range.start.row),
		slotCol: colSlots.indexOf(page.range.start.col),
		range: page.range,
		rows: page.rows,
		columns: page.columns,
	}));
	const empty = !entries.length;
	if (empty)
		entries.push({ number: 1, slotRow: 0, slotCol: 0, range: undefined, rows: [], columns: [] });
	// Excel fills the last page of each axis with the following blank rows and columns.
	const lastRow = Math.max(...entries.map((e) => e.rows.at(-1) ?? -1));
	const lastCol = Math.max(...entries.map((e) => e.columns.at(-1) ?? -1));
	for (const entry of entries) {
		if ((entry.rows.at(-1) ?? -1) === lastRow)
			entry.rows = fillTo(
				entry.rows,
				metrics.rowHeight,
				metrics.isRowHidden,
				plan.contentHeight / plan.scale,
				MAX_ROW,
			);
		if ((entry.columns.at(-1) ?? -1) === lastCol)
			entry.columns = fillTo(
				entry.columns,
				metrics.colWidth,
				metrics.isColHidden,
				plan.contentWidth / plan.scale,
				MAX_COL,
			);
	}
	const side = margins.left * 96;
	const alignedWidth = paperWidth - (margins.left + margins.right) * 96;
	const band = (margin: number, from: number) => Math.max(0, (margin - from) * 96);
	const pages = entries.map((entry): PageViewPage => {
		const x = entry.slotCol * (paperWidth + gap);
		const y = entry.slotRow * (paperHeight + gap);
		return {
			...entry,
			paper: { x, y, width: paperWidth, height: paperHeight },
			content: {
				x: x + side,
				y: y + margins.top * 96,
				width: plan.contentWidth,
				height: plan.contentHeight,
			},
			header: {
				x: x + side,
				y: y + margins.header * 96,
				width: alignedWidth,
				height: band(margins.top, margins.header),
			},
			footer: {
				x: x + side,
				y: y + paperHeight - margins.bottom * 96,
				width: alignedWidth,
				height: band(margins.bottom, margins.footer),
			},
			rowTops: offsets(entry.rows, metrics.rowHeight),
			colLefts: offsets(entry.columns, metrics.colWidth),
			blank: empty,
		};
	});
	return {
		pages,
		scale: plan.scale,
		plan,
		width: (Math.max(...pages.map((p) => p.slotCol)) + 1) * (paperWidth + gap) - gap,
		height: (Math.max(...pages.map((p) => p.slotRow)) + 1) * (paperHeight + gap) - gap,
	};
}

export type PageViewHit =
	| { kind: 'cell'; page: number; row: number; col: number }
	| { kind: 'header' | 'footer'; page: number; section: HeaderFooterSection }
	| { kind: 'blank'; page: number }
	| { kind: 'margin'; page: number };

const inside = (r: PageViewRect, x: number, y: number) =>
	x >= r.x && x < r.x + r.width && y >= r.y && y < r.y + r.height;

/** What lies under canvas point (x, y): a cell, a header or footer section, or the margin. */
export function pageViewHit(
	layout: PageViewLayout,
	sizes: Sizes,
	x: number,
	y: number,
): PageViewHit | undefined {
	const page = layout.pages.find((p) => inside(p.paper, x, y));
	if (!page) return undefined;
	for (const kind of ['header', 'footer'] as const) {
		const area = page[kind];
		if (inside(area, x, y)) {
			const third = (x - area.x) / (area.width / 3);
			return {
				kind,
				page: page.number,
				section: third < 1 ? 'left' : third < 2 ? 'center' : 'right',
			};
		}
	}
	if (!inside(page.content, x, y)) return { kind: 'margin', page: page.number };
	const find = (tops: number[], size: (i: number) => number, indices: number[], at: number) => {
		for (let i = tops.length - 1; i >= 0; i--)
			if (at >= tops[i]!) return at < tops[i]! + size(indices[i]!) ? indices[i] : undefined;
		return undefined;
	};
	const row = find(page.rowTops, sizes.rowHeight, page.rows, (y - page.content.y) / layout.scale);
	const col = find(
		page.colLefts,
		sizes.colWidth,
		page.columns,
		(x - page.content.x) / layout.scale,
	);
	return row === undefined || col === undefined
		? { kind: 'margin', page: page.number }
		: { kind: 'cell', page: page.number, row, col };
}

/** The canvas rectangle of a cell on a page, or undefined when the page does not show it. */
export function pageViewCellRect(
	layout: PageViewLayout,
	sizes: Sizes,
	pageNumber: number,
	row: number,
	col: number,
): PageViewRect | undefined {
	const page = layout.pages.find((p) => p.number === pageNumber);
	const r = page?.rows.indexOf(row) ?? -1;
	const c = page?.columns.indexOf(col) ?? -1;
	if (!page || r < 0 || c < 0) return undefined;
	return {
		x: page.content.x + page.colLefts[c]! * layout.scale,
		y: page.content.y + page.rowTops[r]! * layout.scale,
		width: sizes.colWidth(col) * layout.scale,
		height: sizes.rowHeight(row) * layout.scale,
	};
}

/** The first page that shows a cell (repeated title rows appear on several pages). */
export function pageOfCell(layout: PageViewLayout, row: number, col: number): number | undefined {
	return layout.pages.find((p) => p.rows.includes(row) && p.columns.includes(col))?.number;
}
