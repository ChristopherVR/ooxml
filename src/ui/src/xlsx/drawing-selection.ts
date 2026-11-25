import type { EditorContext } from 'ooxml-core/xlsx/ui';

export function selectedDrawings(ctx: EditorContext): number[] {
	const selection = ctx.selection.get();
	if (selection.sheet !== ctx.activeSheet() || selection.drawing === undefined) return [];
	return selection.drawings ?? [selection.drawing];
}

export function setDrawingSelection(ctx: EditorContext, indices: number[]): void {
	ctx.selection.set({ drawing: indices.at(-1), drawings: indices });
}

export function selectDrawing(ctx: EditorContext, index: number, toggle: boolean): void {
	const current = toggle ? selectedDrawings(ctx) : [];
	setDrawingSelection(
		ctx,
		current.includes(index) ? current.filter((item) => item !== index) : [...current, index],
	);
}
