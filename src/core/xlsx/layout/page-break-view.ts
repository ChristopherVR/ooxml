import type { CellRange } from '../address';
import type { Workbook } from '../model';
import { paginateWorksheet, printRange } from './print-layout';
import { sheetPageBreaks, type PrintAxis } from './print-breaks';

export interface PageBreakLine {
	axis: PrintAxis;
	/** The break lies before this row (`axis: 'row'`) or column. */
	before: number;
	/** Manual breaks are solid in Excel, automatic ones dashed. */
	manual: boolean;
}

export interface PageBreakPage {
	number: number;
	range: CellRange;
}

export interface PageBreakPreview {
	/** The print area, or the used range when none is set. */
	area: CellRange | undefined;
	pages: PageBreakPage[];
	lines: PageBreakLine[];
	/** Fit to pages is on: Excel then places automatic breaks itself and ignores manual ones. */
	fit: boolean;
}

/** Pages, page-break lines and the print area that Page Break Preview paints. */
export function pageBreakPreview(workbook: Workbook, sheetIndex: number): PageBreakPreview {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) throw new RangeError('Invalid worksheet.');
	const setup = sheet.pageSetup ?? {};
	const fit = setup.fitToWidth !== undefined || setup.fitToHeight !== undefined;
	const area = printRange(workbook, sheetIndex);
	const plan = paginateWorksheet(workbook, sheetIndex);
	const pages = plan.pages.map((page) => ({ number: page.number, range: page.range }));
	const lines: PageBreakLine[] = [];
	for (const axis of ['row', 'col'] as const) {
		const manual = new Set(
			fit
				? []
				: sheetPageBreaks(sheet, axis)
						.filter((entry) => entry.manual)
						.map((entry) => entry.before),
		);
		const starts = [...new Set(pages.map((page) => page.range.start[axis]))].sort((a, b) => a - b);
		for (let i = 1; i < starts.length; i++) {
			const start = starts[i]!;
			const previous = starts[i - 1]!;
			const isManual = [...manual].some((id) => id > previous && id <= start);
			lines.push({ axis, before: start, manual: isManual });
		}
	}
	return { area, pages, lines, fit };
}

export interface BreakEdit {
	axis: PrintAxis;
	before: number;
	present: boolean;
}

/**
 * The edits that move a page-break line to sit before `toBefore`: a manual break is removed and
 * re-inserted, an automatic one becomes a manual break at the new place. Nothing when the target
 * is the same, is not inside the print area, or already holds a manual break.
 */
export function planBreakDrag(
	preview: PageBreakPreview,
	line: PageBreakLine,
	toBefore: number,
): BreakEdit[] {
	const { area } = preview;
	if (preview.fit || !area || toBefore === line.before) return [];
	const lo = (line.axis === 'row' ? area.start.row : area.start.col) + 1;
	const hi = line.axis === 'row' ? area.end.row : area.end.col;
	if (!Number.isInteger(toBefore) || toBefore < lo || toBefore > hi) return [];
	if (preview.lines.some((l) => l.axis === line.axis && l.manual && l.before === toBefore))
		return [];
	const edits: BreakEdit[] = [];
	if (line.manual) edits.push({ axis: line.axis, before: line.before, present: false });
	edits.push({ axis: line.axis, before: toBefore, present: true });
	return edits;
}

export type PrintAreaEdge = 'top' | 'bottom' | 'left' | 'right';

/** The print area after dragging one edge so that it includes row or column `to`. */
export function resizePrintArea(
	area: CellRange,
	edge: PrintAreaEdge,
	to: number,
): CellRange | undefined {
	if (!Number.isInteger(to) || to < 0) return undefined;
	const next: CellRange = { start: { ...area.start }, end: { ...area.end } };
	if (edge === 'top') next.start.row = Math.min(to, area.end.row);
	else if (edge === 'bottom') next.end.row = Math.max(to, area.start.row);
	else if (edge === 'left') next.start.col = Math.min(to, area.end.col);
	else next.end.col = Math.max(to, area.start.col);
	return next;
}
