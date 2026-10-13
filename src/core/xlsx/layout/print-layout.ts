import { normalizeRange, type CellRange } from '../address';
import { usedRange } from '../cells';
import type { Workbook } from '../model';
import { createGridMetrics } from './metrics';
import { sheetPageBreaks } from './print-breaks';
import { printRepeat } from './print-repeat';
import { parseHeaderFooter } from './header-footer';

export const PRINT_PAPER_INCHES: Readonly<Record<number, readonly [number, number]>> = {
	1: [8.5, 11],
	5: [8.5, 14],
	7: [7.25, 10.5],
	8: [297 / 25.4, 420 / 25.4],
	9: [210 / 25.4, 297 / 25.4],
	11: [148 / 25.4, 210 / 25.4],
	13: [182 / 25.4, 257 / 25.4],
	20: [4.125, 9.5],
	27: [110 / 25.4, 220 / 25.4],
};
export const PRINT_MARGINS = {
	left: 0.7,
	right: 0.7,
	top: 0.75,
	bottom: 0.75,
	header: 0.3,
	footer: 0.3,
};
export interface PrintPage {
	number: number;
	range: CellRange;
	rows: number[];
	columns: number[];
	width: number;
	height: number;
}
export interface PrintLayout {
	pages: PrintPage[];
	paperWidth: number;
	paperHeight: number;
	contentWidth: number;
	contentHeight: number;
	margins: typeof PRINT_MARGINS;
	scale: number;
	warnings: string[];
}

/** Single rectangular print area, else the used range expanded to cover anchored merges. */
export function printRange(workbook: Workbook, sheetIndex: number): CellRange | undefined {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return undefined;
	if (sheet.pageSetup?.printArea) return normalizeRange(sheet.pageSetup.printArea);
	const used = usedRange(sheet);
	if (!used) return undefined;
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

function chunks(
	indices: number[],
	size: (index: number) => number,
	capacity: number,
	breaks: Set<number>,
): number[][] {
	const sorted = [...breaks].sort((a, b) => a - b);
	let cursor = 0;
	const out: number[][] = [];
	let line: number[] = [];
	let extent = 0;
	for (const index of indices) {
		let manual = false;
		while (cursor < sorted.length && sorted[cursor]! <= index) {
			cursor++;
			manual = true;
		}
		if (line.length && (manual || extent + size(index) > capacity + 0.001)) {
			out.push(line);
			line = [];
			extent = 0;
		}
		line.push(index);
		extent += size(index);
	}
	if (line.length) out.push(line);
	return out;
}

/** Printer-independent pagination in CSS pixels; fonts and device margins may differ in Excel. */
export function paginateWorksheet(workbook: Workbook, sheetIndex: number): PrintLayout {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) throw new RangeError('Invalid worksheet.');
	const setup = sheet.pageSetup ?? {};
	if (
		setup.firstPageNumber !== undefined &&
		(!Number.isInteger(setup.firstPageNumber) ||
			setup.firstPageNumber < 0 ||
			setup.firstPageNumber > 32767)
	)
		throw new Error('Enter a first page number between 0 and 32767.');
	const warnings: string[] = [];
	if (
		!setup.printArea &&
		workbook.definedNames.some(
			(name) => name.name === '_xlnm.Print_Area' && name.localSheet === sheetIndex,
		)
	)
		warnings.push(
			'Multiple or unsupported print areas are preserved; preview uses the worksheet used range.',
		);
	const paper = PRINT_PAPER_INCHES[setup.paperSize ?? 1] ?? PRINT_PAPER_INCHES[1]!;
	if (!PRINT_PAPER_INCHES[setup.paperSize ?? 1])
		warnings.push('This paper size is not supported in preview; Letter is used.');
	const landscape = setup.orientation === 'landscape';
	const paperWidth = paper[landscape ? 1 : 0] * 96;
	const paperHeight = paper[landscape ? 0 : 1] * 96;
	const margins = { ...PRINT_MARGINS, ...setup.margins };
	const contentWidth = paperWidth - (margins.left + margins.right) * 96;
	const contentHeight = paperHeight - (margins.top + margins.bottom) * 96;
	if (
		!Object.values(margins).every((value) => Number.isFinite(value) && value >= 0) ||
		contentWidth <= 0 ||
		contentHeight <= 0
	)
		throw new Error('Margins do not fit the paper size.');
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const range = printRange(workbook, sheetIndex);
	const repeat = printRepeat(workbook, sheetIndex, range, metrics);
	warnings.push(...repeat.warnings);
	const titleRows = new Set(repeat.rows);
	const titleColumns = new Set(repeat.columns);
	const rows: number[] = [];
	const columns: number[] = [];
	if (range) {
		if ((range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1) > 1_000_000)
			throw new Error('Print preview supports at most one million cells.');
		for (let row = range.start.row; row <= range.end.row; row++)
			if (!metrics.isRowHidden(row) && !titleRows.has(row)) rows.push(row);
		for (let col = range.start.col; col <= range.end.col; col++)
			if (!metrics.isColHidden(col) && !titleColumns.has(col)) columns.push(col);
	}
	const headings = sheet.printOptions?.headings ? 24 : 0;
	const largestWidth = columns.reduce((max, col) => Math.max(max, metrics.colWidth(col)), 0);
	const largestHeight = rows.reduce((max, row) => Math.max(max, metrics.rowHeight(row)), 0);
	const fit = setup.fitToWidth !== undefined || setup.fitToHeight !== undefined;
	if (
		setup.scale !== undefined &&
		(!Number.isFinite(setup.scale) || setup.scale < 10 || setup.scale > 400)
	)
		throw new Error('Enter a scale between 10 and 400.');
	let scale = Math.min(4, Math.max(0.1, (setup.scale ?? 100) / 100));
	if (fit) {
		const wide = setup.fitToWidth ?? 1;
		const tall = setup.fitToHeight ?? 1;
		if (![wide, tall].every((value) => Number.isInteger(value) && value >= 0 && value <= 32767))
			throw new Error('Enter a whole number of pages.');
		const fits = (candidate: number) =>
			(largestWidth + repeat.width + headings) * candidate <= contentWidth &&
			(largestHeight + repeat.height + headings) * candidate <= contentHeight &&
			(!wide ||
				chunks(
					columns,
					metrics.colWidth,
					contentWidth / candidate - headings - repeat.width,
					new Set(),
				).length <= wide) &&
			(!tall ||
				chunks(
					rows,
					metrics.rowHeight,
					contentHeight / candidate - headings - repeat.height,
					new Set(),
				).length <= tall);
		let low = 0.1;
		let high = 1;
		if (!fits(low))
			warnings.push('The minimum scale prevents fitting the requested number of pages.');
		else if (fits(high)) low = high;
		else
			for (let i = 0; i < 24; i++) {
				const middle = (low + high) / 2;
				if (fits(middle)) low = middle;
				else high = middle;
			}
		scale = low;
	}
	const breaks = (axis: 'row' | 'col') =>
		new Set(
			sheetPageBreaks(sheet, axis)
				.filter(
					(entry) =>
						entry.manual &&
						!fit &&
						entry.min === 0 &&
						entry.max === (axis === 'row' ? 16383 : 1048575),
				)
				.map((entry) => entry.before),
		);
	if (
		fit &&
		[...sheetPageBreaks(sheet, 'row'), ...sheetPageBreaks(sheet, 'col')].some(
			(entry) => entry.manual,
		)
	)
		warnings.push('Fit to pages ignores manual page breaks.');
	const rowChunks = chunks(
		rows,
		metrics.rowHeight,
		contentHeight / scale - headings - repeat.height,
		breaks('row'),
	);
	const colChunks = chunks(
		columns,
		metrics.colWidth,
		contentWidth / scale - headings - repeat.width,
		breaks('col'),
	);
	if (
		(repeat.width + headings) * scale >= contentWidth ||
		(repeat.height + headings) * scale >= contentHeight
	)
		throw new Error('Repeated titles leave no room for worksheet data.');
	if (!rowChunks.length && repeat.rows.length) rowChunks.push([]);
	if (!colChunks.length && repeat.columns.length) colChunks.push([]);
	if (
		rows.some((row) => (metrics.rowHeight(row) + headings) * scale > contentHeight) ||
		columns.some((col) => (metrics.colWidth(col) + headings) * scale > contentWidth)
	)
		warnings.push('Rows or columns larger than a printed page are clipped.');
	if (sheet.drawings.length) warnings.push('Preview does not render drawings.');
	if (
		[
			setup.header,
			setup.footer,
			setup.evenHeader,
			setup.evenFooter,
			setup.firstHeader,
			setup.firstFooter,
		].some((value) => value && parseHeaderFooter(value).unsupported.length)
	)
		warnings.push('Unsupported header/footer codes are preserved but not rendered.');
	if (
		(['row', 'col'] as const).some((axis) =>
			sheetPageBreaks(sheet, axis).some(
				(entry) =>
					entry.manual && (entry.min !== 0 || entry.max !== (axis === 'row' ? 16383 : 1048575)),
			),
		)
	)
		warnings.push('Partial page breaks are preserved but not used in preview.');
	const pages: PrintPage[] = [];
	if (rowChunks.length * colChunks.length > 1000)
		throw new Error('Print preview supports at most 1000 pages.');
	const add = (r: number[], c: number[]) =>
		pages.push({
			number: pages.length + 1,
			range: {
				start: { row: r[0] ?? repeat.rows[0]!, col: c[0] ?? repeat.columns[0]! },
				end: { row: r.at(-1) ?? repeat.rows.at(-1)!, col: c.at(-1) ?? repeat.columns.at(-1)! },
			},
			rows: [...repeat.rows, ...r],
			columns: [...repeat.columns, ...c],
			width: repeat.width + c.reduce((sum, col) => sum + metrics.colWidth(col), 0),
			height: repeat.height + r.reduce((sum, row) => sum + metrics.rowHeight(row), 0),
		});
	if (setup.pageOrder === 'overThenDown')
		for (const r of rowChunks) for (const c of colChunks) add(r, c);
	else for (const c of colChunks) for (const r of rowChunks) add(r, c);
	if (pages.reduce((count, page) => count + page.rows.length * page.columns.length, 0) > 1_000_000)
		throw new Error('Print preview supports at most one million cells.');
	if (
		sheet.merges.some((merge) =>
			pages.some(
				(page) =>
					page.range.start.row <= merge.end.row &&
					page.range.end.row >= merge.start.row &&
					page.range.start.col <= merge.end.col &&
					page.range.end.col >= merge.start.col &&
					(merge.start.row < page.range.start.row ||
						merge.end.row > page.range.end.row ||
						merge.start.col < page.range.start.col ||
						merge.end.col > page.range.end.col),
			),
		)
	)
		warnings.push('Merged cells crossing pages are rendered separately on each page.');
	return { pages, paperWidth, paperHeight, contentWidth, contentHeight, margins, scale, warnings };
}
