import type { ColumnInfo, RowInfo } from '../model.js';
import { editColumns } from './columns.js';
import { type EditContext, sheetAt } from './context.js';

/** Excel's deepest outline level. */
export const MAX_OUTLINE_LEVEL = 7;

function nextLevel(level: number | undefined, delta: 1 | -1): number {
	return Math.max(0, Math.min(MAX_OUTLINE_LEVEL, (level ?? 0) + delta));
}

function span(from: number, to: number): number[] {
	const lo = Math.max(0, Math.min(from, to));
	const hi = Math.max(from, to);
	return Array.from({ length: hi - lo + 1 }, (_v, i) => lo + i);
}

function outlineRows(ctx: EditContext, s: number, from: number, to: number, delta: 1 | -1): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		delta > 0 ? 'Group rows' : 'Ungroup rows',
		'structure',
		[{ kind: 'sheet', sheet: s }],
		() => {
			for (const row of span(from, to)) {
				const info: RowInfo = { ...sheet.rowInfo.get(row) };
				const level = nextLevel(info.outlineLevel, delta);
				if (level) info.outlineLevel = level;
				else {
					delete info.outlineLevel;
					delete info.collapsed;
				}
				if (Object.keys(info).length) sheet.rowInfo.set(row, info);
				else sheet.rowInfo.delete(row);
			}
		},
		{ sheet: s, structural: true },
	);
}

function outlineColumns(
	ctx: EditContext,
	s: number,
	from: number,
	to: number,
	delta: 1 | -1,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		delta > 0 ? 'Group columns' : 'Ungroup columns',
		'structure',
		[{ kind: 'sheet', sheet: s }],
		() =>
			editColumns(sheet, span(from, to), (info: ColumnInfo) => {
				const level = nextLevel(info.outlineLevel, delta);
				if (level) info.outlineLevel = level;
				else {
					delete info.outlineLevel;
					delete info.collapsed;
				}
			}),
		{ sheet: s, structural: true },
	);
}

/** Raises the outline level of rows `from`..`to` (inclusive) by one, up to 7. */
export const groupRows = (ctx: EditContext, s: number, from: number, to: number): void =>
	outlineRows(ctx, s, from, to, 1);
/** Lowers the outline level of rows `from`..`to` (inclusive) by one. */
export const ungroupRows = (ctx: EditContext, s: number, from: number, to: number): void =>
	outlineRows(ctx, s, from, to, -1);
/** Raises the outline level of columns `from`..`to` (inclusive) by one, up to 7. */
export const groupColumns = (ctx: EditContext, s: number, from: number, to: number): void =>
	outlineColumns(ctx, s, from, to, 1);
/** Lowers the outline level of columns `from`..`to` (inclusive) by one. */
export const ungroupColumns = (ctx: EditContext, s: number, from: number, to: number): void =>
	outlineColumns(ctx, s, from, to, -1);

/** Sets the sheet's standard column width in characters (0-255); undefined restores the default. */
export function setDefaultColumnWidth(
	ctx: EditContext,
	s: number,
	width: number | undefined,
): void {
	if (width !== undefined && (!Number.isFinite(width) || width < 0 || width > 255))
		throw new RangeError('The standard column width must be between 0 and 255 characters.');
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Standard width',
		'format',
		[{ kind: 'sheet', sheet: s }],
		() => {
			if (width === undefined) delete sheet.defaultColWidth;
			else sheet.defaultColWidth = width;
		},
		{ sheet: s, structural: true },
	);
}
