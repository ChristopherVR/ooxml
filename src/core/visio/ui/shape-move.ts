import type { VisioEdit } from '../edit-commands';
import type { VisioMatrix, VisioPage, VisioShape } from '../model';
import { visioStyleFormattingShape } from './formatting';
import { visioPageEditToDrawing } from './page-edit';
import { visioMovableGroup } from './shape-group';

export interface VisioPagePoint {
	x: number;
	y: number;
}
/** Source admission remains authoritative for locks, formulas, geometry and dependencies. */
export function visioMovementShape(page: VisioPage, id: string): VisioShape | undefined {
	const shape = visioStyleFormattingShape(page, id) ?? visioMovableGroup(page, id);
	if (
		!shape ||
		(shape.kind !== 'shape' && !(shape.kind === 'group' && shape.children.length)) ||
		!shape.rotation ||
		!(shape.width > 0 && shape.height > 0) ||
		![
			shape.width,
			shape.height,
			shape.rotation.pinX,
			shape.rotation.pinY,
			...shape.transform,
		].every(Number.isFinite) ||
		page.connectors.some(
			(connection) => connection.fromShapeId === id || connection.toShapeId === id,
		)
	)
		return undefined;
	return shape;
}
/** Client points have been converted to physical SVG page inches; commands use upward y. */
export function visioPageDragDelta(start: VisioPagePoint, end: VisioPagePoint): VisioPagePoint {
	return { x: end.x - start.x, y: start.y - end.y };
}
export function visioMovePreviewTransform(
	shape: VisioShape,
	delta: VisioPagePoint,
): VisioMatrix | undefined {
	if (![delta.x, delta.y, ...shape.transform].every(Number.isFinite)) return undefined;
	const [a, b, c, d, e, f] = shape.transform;
	const result: VisioMatrix = [a, b, c, d, e + delta.x, f + delta.y];
	return result.every(Number.isFinite) ? result : undefined;
}
/** One common translation over the entire eligible selection; callers commit it atomically. */
export function visioMoveCommands(
	page: VisioPage,
	ids: readonly string[],
	delta: VisioPagePoint,
): VisioEdit[] | undefined {
	if (
		!ids.length ||
		ids.length > 1000 ||
		new Set(ids).size !== ids.length ||
		![delta.x, delta.y].every(Number.isFinite)
	)
		return undefined;
	const shapes = ids.map((id) => visioMovementShape(page, id));
	if (shapes.some((shape) => !shape)) return undefined;
	if (delta.x === 0 && delta.y === 0) return [];
	try {
		const edits = shapes.map((shape) =>
			visioPageEditToDrawing(page, {
				type: 'move-shape',
				pageId: page.id,
				shapeId: shape!.id,
				x: shape!.rotation!.pinX + delta.x,
				y: shape!.rotation!.pinY + delta.y,
			}),
		);
		return edits.every(
			(edit) => edit.type === 'move-shape' && Number.isFinite(edit.x) && Number.isFinite(edit.y),
		)
			? edits
			: undefined;
	} catch {
		return undefined;
	}
}
