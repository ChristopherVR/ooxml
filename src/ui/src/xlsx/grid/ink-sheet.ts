// The sheet's ink as the Draw tools see it: each drawing with its box at 100% zoom, the anchor of a
// box, and the change that keeps a drawing's box fitted to the strokes it keeps.
import type { InkBox, InkStroke } from 'ooxml-core/drawingml';
import { anchorToPixelBox, pixelBoxToAnchor, type InkObject } from 'ooxml-core/xlsx';
import { inkTools } from '../../ink/tools';
import type { GridView } from './grid-view';

export interface SheetInk {
	index: number;
	ink: InkObject;
	/** Its box on the sheet, in pixels at 100%. */
	box: InkBox;
}

/** Ink coordinates are kept at 100%; the grid's plane is drawn at the sheet's zoom. */
export const zoomOf = (view: GridView): number => view.zoom() / 100;

/**
 * Every ink drawing of the shown sheet with its box at 100%. `editable` leaves out ink read only
 * in part (`InkObject.truncated`), whose strokes the tools must not change.
 */
export function sheetInks(view: GridView, editable = true): SheetInk[] {
	const sheet = view.sheet();
	if (!sheet) return [];
	const f = zoomOf(view);
	return sheet.drawings.flatMap((drawing, index) => {
		if (drawing.kind !== 'ink' || (editable && drawing.truncated)) return [];
		const box = anchorToPixelBox(sheet, drawing.anchor, view.metrics);
		return [
			{
				index,
				ink: drawing,
				box: { x: box.x / f, y: box.y / f, width: box.w / f, height: box.h / f },
			},
		];
	});
}

/** The anchor of a box on the sheet (pixels at 100%). */
export function anchorOf(view: GridView, box: InkBox) {
	const sheet = view.sheet()!;
	const f = zoomOf(view);
	return pixelBoxToAnchor(
		sheet,
		{ x: box.x * f, y: box.y * f, w: box.width * f, h: box.height * f },
		view.metrics,
	);
}

/**
 * The drawing `entry` keeping only `strokes` (its own coordinates): its anchor fitted to them, as a
 * new stroke fits a growing one, so erasing its edge never leaves a frame larger than its ink.
 */
export function fittedInk(view: GridView, entry: SheetInk, strokes: InkStroke[]) {
	const { ink, box } = entry;
	const fit = inkTools().fitInkBox(strokes, ink, box);
	return fit
		? {
				anchor: anchorOf(view, fit.box),
				ink: { strokes: fit.strokes, width: fit.width, height: fit.height },
			}
		: { anchor: ink.anchor, ink: { strokes, width: ink.width, height: ink.height } };
}
