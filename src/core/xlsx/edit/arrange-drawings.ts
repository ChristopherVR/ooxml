import { anchorToPixelBox, pixelBoxToAnchor } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { DrawingObject } from '../model';
import { type EditContext, sheetAt } from './context';
import { editDrawings, isWorksheetShape } from './native-shapes';

export type DrawingAlignment = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';

export function reorderDrawing(
	ctx: EditContext,
	s: number,
	index: number,
	direction: -1 | 1,
): number {
	const sheet = sheetAt(ctx.workbook, s);
	const drawing = sheet.drawings[index];
	const destination = index + direction;
	if (!drawing || !sheet.drawings[destination]) return index;
	return editDrawings(ctx, s, direction === 1 ? 'Bring forward' : 'Send backward', () => {
		sheet.drawings.splice(index, 1);
		sheet.drawings.splice(destination, 0, drawing);
		return destination;
	});
}

export function alignDrawings(
	ctx: EditContext,
	s: number,
	indices: readonly number[],
	alignment: DrawingAlignment,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const items = [...new Set(indices)].flatMap((index) => {
		const drawing = sheet.drawings[index];
		return drawing && drawing.kind !== 'formControl'
			? [{ drawing, box: anchorToPixelBox(sheet, drawing.anchor, metrics) }]
			: [];
	});
	if (items.length < 2) return;
	const horizontal = alignment === 'left' || alignment === 'center' || alignment === 'right';
	const axis = horizontal ? 'x' : 'y';
	const size = horizontal ? 'w' : 'h';
	const start = Math.min(...items.map(({ box }) => box[axis]));
	const end = Math.max(...items.map(({ box }) => box[axis] + box[size]));
	const position = (span: number) => {
		if (alignment === 'left' || alignment === 'top') return start;
		if (alignment === 'right' || alignment === 'bottom') return end - span;
		return (start + end - span) / 2;
	};
	editDrawings(ctx, s, 'Align objects', () => {
		for (const { drawing, box } of items) {
			const point = position(box[size]);
			if (box[axis] !== point)
				drawing.anchor = pixelBoxToAnchor(
					sheet,
					{ ...box, [axis]: point },
					metrics,
					drawing.anchor,
				);
		}
	});
}

export function canRotateDrawing(drawing: DrawingObject | undefined): boolean {
	return isWorksheetShape(drawing) || drawing?.kind === 'image';
}

export function rotateDrawing(ctx: EditContext, s: number, index: number, degrees: number): void {
	const drawing = sheetAt(ctx.workbook, s).drawings[index];
	if (!drawing || !Number.isFinite(degrees)) return;
	const rotated = (old: number) => (((old + degrees) % 360) + 360) % 360;
	if (isWorksheetShape(drawing)) {
		editDrawings(ctx, s, 'Rotate object', () => {
			drawing.shape.rotation = rotated(drawing.shape.rotation ?? 0);
			if (drawing.sourceXml) drawing.shapeEdit = { ...drawing.shapeEdit, transform: true };
		});
	} else if (drawing.kind === 'image') {
		editDrawings(ctx, s, 'Rotate object', () => {
			drawing.rotation = rotated(drawing.rotation ?? 0);
		});
	}
}

export function flipDrawing(
	ctx: EditContext,
	s: number,
	index: number,
	axis: 'horizontal' | 'vertical',
): void {
	const drawing = sheetAt(ctx.workbook, s).drawings[index];
	const field = axis === 'horizontal' ? 'flipH' : 'flipV';
	if (isWorksheetShape(drawing)) {
		editDrawings(ctx, s, 'Flip object', () => {
			drawing.shape[field] = !drawing.shape[field];
			if (drawing.sourceXml) drawing.shapeEdit = { ...drawing.shapeEdit, transform: true };
		});
	} else if (drawing?.kind === 'image') {
		editDrawings(ctx, s, 'Flip object', () => {
			drawing[field] = !drawing[field];
		});
	}
}

export function drawingHidden(drawing: DrawingObject): boolean {
	return !!(drawing.kind === 'image'
		? drawing.hidden
		: isWorksheetShape(drawing) && drawing.shape.hidden);
}

export function setDrawingHidden(
	ctx: EditContext,
	s: number,
	index: number,
	hidden: boolean,
): void {
	const drawing = sheetAt(ctx.workbook, s).drawings[index];
	if (!drawing || drawingHidden(drawing) === hidden) return;
	if (isWorksheetShape(drawing)) {
		editDrawings(ctx, s, hidden ? 'Hide object' : 'Show object', () => {
			drawing.shape.hidden = hidden;
			if (drawing.sourceXml) drawing.shapeEdit = { ...drawing.shapeEdit, visibility: true };
		});
	} else if (drawing.kind === 'image') {
		editDrawings(ctx, s, hidden ? 'Hide object' : 'Show object', () => {
			drawing.hidden = hidden;
		});
	}
}
