import type { DrawingAnchor, DrawingObject } from '../model.js';
import { type EditContext, sheetAt } from './context.js';

function drawingAt(ctx: EditContext, s: number, index: number): DrawingObject {
	const drawing = sheetAt(ctx.workbook, s).drawings[index];
	if (!drawing) throw new RangeError(`No drawing at index ${index} on sheet ${s}`);
	return drawing;
}

function checkAnchor(anchor: DrawingAnchor): void {
	const points = [anchor.from, ...(anchor.to ? [anchor.to] : [])];
	for (const p of points)
		if (
			![p.row, p.col, p.rowOffset, p.colOffset].every(Number.isFinite) ||
			p.row < 0 ||
			p.col < 0 ||
			p.rowOffset < 0 ||
			p.colOffset < 0
		)
			throw new RangeError('A drawing anchor needs non-negative cell positions and offsets.');
	if (anchor.ext && !(anchor.ext.cx >= 0 && anchor.ext.cy >= 0))
		throw new RangeError('A drawing extent cannot be negative.');
}

/**
 * Moves or resizes a picture, chart or shape (one undo step). The saved drawing part is
 * regenerated from the model, so the new anchor is what Excel shows.
 */
export function setDrawingAnchor(
	ctx: EditContext,
	s: number,
	index: number,
	anchor: DrawingAnchor,
): void {
	const drawing = drawingAt(ctx, s, index);
	checkAnchor(anchor);
	ctx.run(
		'Move object',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			drawing.anchor = structuredClone(anchor);
		},
		{ sheet: s },
	);
}

/** Deletes a picture, chart or shape (one undo step). */
export function deleteDrawing(ctx: EditContext, s: number, index: number): void {
	drawingAt(ctx, s, index);
	const sheet = sheetAt(ctx.workbook, s);
	ctx.run(
		'Delete object',
		'annotations',
		[{ kind: 'sheet', sheet: s }],
		() => {
			sheet.drawings.splice(index, 1);
		},
		{ sheet: s },
	);
}
