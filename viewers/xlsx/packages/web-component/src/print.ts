/**
 * Printing: the active sheet's print area (or its used range) rendered as a print-only HTML table
 * from the core's cell views, column and row metrics, and merges, printed through a hidden frame.
 * It approximates Excel's page output (no headers, footers or scaling to fit) and says so.
 */
import {
	cellView,
	createGridMetrics,
	mergeView,
	usedRange,
	type CellRange,
	type CellView,
	type EdgeView,
	type Workbook,
} from '@christophervr/xlsx-core';

const escapeHtml = (text: string) => text.replace(/[&<>"']/g, (ch) => `&#${ch.charCodeAt(0)};`);

const edge = (view: EdgeView | undefined) =>
	view ? `${view.style === 'double' ? 3 : view.widthPx}px ${view.style} ${view.color}` : 'none';

/** Inline CSS for one cell view. */
export function cellCss(view: CellView): string {
	const { font } = view;
	const parts = [
		`font-family:"${font.family.replace(/["\\]/g, '')}",Calibri,sans-serif`,
		`font-size:${font.sizePx}px`,
		`color:${font.color}`,
		`text-align:${view.hAlign === 'centerContinuous' ? 'center' : view.hAlign === 'distributed' || view.hAlign === 'fill' ? 'justify' : view.hAlign}`,
		`vertical-align:${view.vAlign === 'center' ? 'middle' : view.vAlign === 'top' ? 'top' : 'bottom'}`,
		`white-space:${view.wrap ? 'pre-wrap' : 'pre'}`,
		`border-top:${edge(view.borders.top)}`,
		`border-right:${edge(view.borders.right)}`,
		`border-bottom:${edge(view.borders.bottom)}`,
		`border-left:${edge(view.borders.left)}`,
	];
	if (font.bold) parts.push('font-weight:700');
	if (font.italic) parts.push('font-style:italic');
	const lines = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(
		Boolean,
	);
	if (lines.length) parts.push(`text-decoration:${lines.join(' ')}`);
	if (view.indentPx) parts.push(`padding-left:${view.indentPx + 2}px`);
	const fill = view.fill;
	if (fill && 'background' in fill) parts.push(`background:${fill.background}`);
	else if (fill && 'gradient' in fill) parts.push(`background:${fill.gradient}`);
	else if (fill && 'pattern' in fill) parts.push(`background:${fill.fg}`);
	return parts.join(';');
}

/** The range a sheet prints: its print area, else its used range (undefined for an empty sheet). */
export function printRange(workbook: Workbook, sheetIndex: number): CellRange | undefined {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return undefined;
	if (sheet.pageSetup?.printArea) return sheet.pageSetup.printArea;
	const used = usedRange(sheet);
	if (!used) return undefined;
	// A merge anchored in the used range prints whole, as Excel prints it.
	const end = { ...used.end };
	for (const merge of sheet.merges)
		if (
			merge.start.row <= end.row &&
			merge.start.col <= end.col &&
			merge.start.row >= used.start.row &&
			merge.start.col >= used.start.col
		) {
			end.row = Math.max(end.row, merge.end.row);
			end.col = Math.max(end.col, merge.end.col);
		}
	return { start: used.start, end };
}

/** A standalone HTML document with the sheet's print range as a table. */
export function buildPrintHtml(workbook: Workbook, sheetIndex: number, title: string): string {
	const sheet = workbook.sheets[sheetIndex];
	const range = printRange(workbook, sheetIndex);
	const head = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
@page{margin:12mm}body{margin:0;font-family:Calibri,sans-serif}table{border-collapse:collapse;table-layout:fixed}
td{overflow:hidden;padding:0 3px;box-sizing:border-box;line-height:1.2}</style></head><body>`;
	if (!sheet || !range) return `${head}</body></html>`;
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const cols: number[] = [];
	for (let c = range.start.col; c <= range.end.col; c++) if (!metrics.isColHidden(c)) cols.push(c);
	const colgroup = cols.map((c) => `<col style="width:${metrics.colWidth(c)}px">`).join('');
	const width = cols.reduce((sum, c) => sum + metrics.colWidth(c), 0);
	const rows: string[] = [];
	for (let r = range.start.row; r <= range.end.row; r++) {
		if (metrics.isRowHidden(r)) continue;
		const cells: string[] = [];
		for (const c of cols) {
			const merge = mergeView(sheet, r, c);
			if (merge.hidden) continue;
			const view = cellView(workbook, sheetIndex, r, c);
			let span = '';
			if (merge.anchor && merge.range) {
				const colspan = cols.filter(
					(col) => col >= merge.range!.start.col && col <= merge.range!.end.col,
				).length;
				let rowspan = 0;
				for (
					let row = merge.range.start.row;
					row <= Math.min(merge.range.end.row, range.end.row);
					row++
				)
					if (!metrics.isRowHidden(row)) rowspan++;
				span = `${colspan > 1 ? ` colspan="${colspan}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''}`;
			}
			cells.push(`<td${span} style="${escapeHtml(cellCss(view))}">${escapeHtml(view.text)}</td>`);
		}
		rows.push(`<tr style="height:${metrics.rowHeight(r)}px">${cells.join('')}</tr>`);
	}
	return `${head}<table style="width:${width}px"><colgroup>${colgroup}</colgroup><tbody>${rows.join('')}</tbody></table></body></html>`;
}

/** How long a print frame may stay without `afterprint` before it is removed anyway. */
export const PRINT_FRAME_TIMEOUT_MS = 60_000;
/** The frame of the previous print, per document, removed when the next print starts. */
const pendingFrames = new WeakMap<Document, () => void>();

/**
 * Prints through a hidden frame so the editor page itself is never reflowed. The frame is removed
 * on `afterprint` (some browsers return from `print()` before the dialog closes, so removing it
 * then would cancel the job), on the next print, or after a long safety timeout.
 */
export function printHtml(doc: Document, html: string): void {
	pendingFrames.get(doc)?.();
	const frame = doc.createElement('iframe');
	frame.setAttribute('aria-hidden', 'true');
	frame.tabIndex = -1;
	frame.style.cssText =
		'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
	doc.body.append(frame);
	const target = frame.contentWindow;
	const frameDoc = frame.contentDocument;
	if (!target || !frameDoc) {
		frame.remove();
		return;
	}
	frameDoc.open();
	frameDoc.write(html);
	frameDoc.close();
	let safety: ReturnType<typeof setTimeout> | undefined;
	const cleanup = () => {
		if (safety !== undefined) clearTimeout(safety);
		target.removeEventListener('afterprint', cleanup);
		frame.remove();
		if (pendingFrames.get(doc) === cleanup) pendingFrames.delete(doc);
	};
	pendingFrames.set(doc, cleanup);
	target.addEventListener('afterprint', cleanup);
	safety = setTimeout(cleanup, PRINT_FRAME_TIMEOUT_MS);
	setTimeout(() => {
		if (!frame.isConnected) return;
		try {
			target.focus();
			target.print();
		} catch (error) {
			// Nothing awaits this timer, so report instead of throwing an uncaught error.
			cleanup();
			console.error('xlsx-editor: printing failed', error);
		}
	}, 50);
}

/** Prints the active sheet of `workbook`. */
export function printWorkbook(
	doc: Document,
	workbook: Workbook,
	sheetIndex: number,
	title: string,
): void {
	printHtml(doc, buildPrintHtml(workbook, sheetIndex, title));
}
