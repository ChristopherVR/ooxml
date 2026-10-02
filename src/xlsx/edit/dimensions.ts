import type { CellStyle, Worksheet } from '../model.js';
import { styleAt } from '../styles.js';
import { editColumns } from './columns.js';
import { type EditContext, displayText, sheetAt } from './context.js';

/** Maximum digit width of the default font (Calibri 11) in pixels. */
const MDW = 7;
/** Excel's widest column, in characters. */
const MAX_WIDTH = 255;
const MAX_HEIGHT = 409;

/**
 * Text width in pixels when no measure is given: a rough per-character estimate scaled by font
 * size (the UI passes a canvas measure for real fitting).
 */
function estimateWidth(text: string, style: CellStyle): number {
	const scale = (style.font.size ?? 11) / 11;
	const longest = Math.max(...text.split('\n').map((line) => line.length));
	return longest * MDW * scale * (style.font.bold ? 1.1 : 1);
}

/** Pixels of text to a column width in characters (Excel adds 5px of padding). */
const pixelsToWidth = (px: number): number => Math.round(((px + 5) / MDW) * 256) / 256;

/** The width that fits a column's content, or undefined when the column is empty. */
export function fitColumnWidth(
	ctx: EditContext,
	sheet: Worksheet,
	col: number,
	measure: (text: string, style: CellStyle) => number,
): number | undefined {
	let widest = -1;
	for (const [row, cells] of sheet.rows) {
		const cell = cells.get(col);
		if (!cell || sheet.rowInfo.get(row)?.hidden) continue;
		const text = displayText(ctx.workbook, cell);
		if (!text) continue;
		widest = Math.max(widest, measure(text, styleAt(ctx.workbook, cell.styleId)));
	}
	return widest < 0 ? undefined : Math.min(MAX_WIDTH, pixelsToWidth(widest));
}

/**
 * Sets column widths in characters. `auto` fits the content using `measure` (text width in
 * pixels); an empty column goes back to the default width.
 */
export function setColumnWidth(
	ctx: EditContext,
	s: number,
	cols: number[],
	width: number | 'auto',
	measure?: (text: string, style: CellStyle) => number,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!cols.length) return;
	const widths = new Map<number, number | undefined>();
	for (const col of cols)
		widths.set(
			col,
			width === 'auto'
				? fitColumnWidth(ctx, sheet, col, measure ?? estimateWidth)
				: Math.max(0, Math.min(MAX_WIDTH, width)),
		);
	ctx.run(
		'Column width',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() =>
			editColumns(sheet, cols, (info) => {
				const next = widths.get(info.min);
				if (next === undefined) {
					delete info.width;
					delete info.customWidth;
					delete info.bestFit;
					return;
				}
				info.width = next;
				info.customWidth = true;
				if (width === 'auto') info.bestFit = true;
				else delete info.bestFit;
				if (next === 0) info.hidden = true;
				else delete info.hidden;
			}),
		{ sheet: s, structural: true },
	);
}

/** The row height that fits the largest font (and wrapped lines) in a row. */
function fitRowHeight(ctx: EditContext, sheet: Worksheet, row: number): number {
	let height = sheet.defaultRowHeight;
	for (const cell of sheet.rows.get(row)?.values() ?? []) {
		const style = styleAt(ctx.workbook, cell.styleId);
		const lineHeight = Math.ceil((style.font.size ?? 11) * (15 / 11) * 4) / 4;
		const lines = style.alignment?.wrapText
			? displayText(ctx.workbook, cell).split('\n').length
			: 1;
		height = Math.max(height, lineHeight * lines);
	}
	return Math.min(MAX_HEIGHT, height);
}

/** Sets row heights in points; `auto` fits the tallest font and clears the custom-height flag. */
export function setRowHeight(
	ctx: EditContext,
	s: number,
	rows: number[],
	height: number | 'auto',
): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!rows.length) return;
	ctx.run(
		'Row height',
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			for (const row of rows) {
				const info = { ...sheet.rowInfo.get(row) };
				if (height === 'auto') {
					const fit = fitRowHeight(ctx, sheet, row);
					if (fit === sheet.defaultRowHeight) delete info.height;
					else info.height = fit;
					delete info.customHeight;
				} else {
					const h = Math.max(0, Math.min(MAX_HEIGHT, height));
					info.height = h;
					info.customHeight = true;
					if (h === 0) info.hidden = true;
					else delete info.hidden;
				}
				if (Object.keys(info).length) sheet.rowInfo.set(row, info);
				else sheet.rowInfo.delete(row);
			}
		},
		{ sheet: s, structural: true },
	);
}

export function setHidden(
	ctx: EditContext,
	s: number,
	axis: 'row' | 'col',
	indices: number[],
	hidden: boolean,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	if (!indices.length) return;
	const noun = axis === 'row' ? 'rows' : 'columns';
	ctx.run(
		`${hidden ? 'Hide' : 'Unhide'} ${noun}`,
		'view',
		[{ kind: 'sheet', sheet: s }],
		() => {
			if (axis === 'col') {
				editColumns(sheet, indices, (info) => {
					if (hidden) info.hidden = true;
					else {
						delete info.hidden;
						if (info.width === 0) delete info.width;
					}
				});
				return;
			}
			for (const row of indices) {
				const info = { ...sheet.rowInfo.get(row) };
				if (hidden) info.hidden = true;
				else {
					delete info.hidden;
					if (info.height === 0) delete info.height;
				}
				if (Object.keys(info).length) sheet.rowInfo.set(row, info);
				else sheet.rowInfo.delete(row);
			}
		},
		{ sheet: s, structural: true },
	);
}
