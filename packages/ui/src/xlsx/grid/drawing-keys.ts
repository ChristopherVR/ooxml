// Keyboard for a selected picture or chart (the grid's focus sink receives the keys): Delete,
// Escape and arrow-key nudging.
import type { EditorContext } from '../context.js';
import type { DrawingLayer } from './drawings.js';

const NUDGE: Record<string, [number, number]> = {
	ArrowUp: [0, -1],
	ArrowDown: [0, 1],
	ArrowLeft: [-1, 0],
	ArrowRight: [1, 0],
};

/**
 * Keys while a picture or chart is selected: Delete removes it, Escape returns to the cells,
 * arrows nudge it. Returns true when the key was handled.
 */
export function drawingKeyDown(
	ctx: EditorContext,
	layer: DrawingLayer,
	event: KeyboardEvent,
): boolean {
	const selection = ctx.selection.get();
	if (selection.drawing === undefined || event.isComposing) return false;
	const handled = () => {
		event.preventDefault();
		event.stopPropagation();
		return true;
	};
	if (event.key === 'Delete' || event.key === 'Backspace') {
		const session = ctx.session();
		if (session && !ctx.readOnly()) {
			ctx.selection.set({ drawing: undefined });
			session.deleteDrawing(selection.sheet, selection.drawing);
		}
		return handled();
	}
	if (event.key === 'Escape') {
		ctx.selection.set({ drawing: undefined });
		return handled();
	}
	const step = NUDGE[event.key];
	if (step && !event.altKey) {
		const by = event.ctrlKey || event.metaKey ? 1 : 4;
		layer.nudge(step[0] * by, step[1] * by);
		return handled();
	}
	return false;
}
