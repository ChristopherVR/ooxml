import type { InkStroke } from '../../drawingml/inkml/types';
import type { DrawingAnchor, InkObject } from '../model';
import { type EditContext, sheetAt, sheetDrawingAt } from './context';
import { checkAnchor } from './drawings';
import { editDrawings } from './native-shapes';

/** Ink for the Draw tab: strokes in a `width` x `height` pixel space stretched onto the anchor. */
export interface NewInk {
	strokes: InkStroke[];
	width: number;
	height: number;
	name?: string;
}

function checkInk(ink: NewInk): void {
	if (!(ink.width > 0 && ink.height > 0) || !ink.strokes.every((stroke) => stroke.path))
		throw new RangeError('Ink needs a size and strokes with paths.');
}

/** Draw > a pen stroke: adds an ink drawing over the cells; returns its drawing index. */
export function addInk(ctx: EditContext, s: number, anchor: DrawingAnchor, ink: NewInk): number {
	checkAnchor(anchor);
	checkInk(ink);
	const sheet = sheetAt(ctx.workbook, s);
	const drawing: InkObject = {
		kind: 'ink',
		anchor: structuredClone(anchor),
		strokes: structuredClone(ink.strokes),
		width: ink.width,
		height: ink.height,
		...(ink.name ? { name: ink.name } : {}),
	};
	return editDrawings(ctx, s, 'Ink', () => sheet.drawings.push(drawing) - 1);
}

/**
 * Replaces the strokes of the ink at `index` (erased, added to or moved by the lasso), with its new
 * box; ink left with no strokes is deleted, as the eraser deletes it in Excel. Edits passing the
 * same `gesture` one after another are one undo step (one eraser stroke).
 */
export function setInk(
	ctx: EditContext,
	s: number,
	index: number,
	anchor: DrawingAnchor,
	ink: NewInk,
	gesture?: string,
): void {
	const drawing = sheetDrawingAt(ctx.workbook, s, index);
	if (drawing?.kind !== 'ink') throw new RangeError(`No ink at index ${index} on sheet ${s}`);
	checkAnchor(anchor);
	if (ink.strokes.length) checkInk(ink);
	editDrawings(
		ctx,
		s,
		ink.strokes.length ? 'Ink' : 'Erase ink',
		() => {
			const sheet = sheetAt(ctx.workbook, s);
			if (!ink.strokes.length) {
				sheet.drawings.splice(index, 1);
				return;
			}
			drawing.anchor = structuredClone(anchor);
			drawing.strokes = structuredClone(ink.strokes);
			drawing.width = ink.width;
			drawing.height = ink.height;
		},
		gesture,
	);
}
