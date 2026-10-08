import {
	alignElements,
	type AlignEdge,
	type BoundingBoxElement,
	type ElementPosition,
} from '../../geometry/align-distribute';
import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioPageEditToDrawing } from './page-edit';
import { visioStyleFormattingShape } from './formatting';

export type VisioArrangement =
	| { type: 'align'; edge: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom' }
	| { type: 'distribute'; axis: 'horizontal' | 'vertical' };

/** Native spacing sorts by centers in upward-positive coordinates, preserving selection ties. */
function distributeBoxes(
	boxes: readonly BoundingBoxElement[],
	axis: 'horizontal' | 'vertical',
): Map<string, ElementPosition> {
	const horizontal = axis === 'horizontal';
	const start = (box: BoundingBoxElement) => (horizontal ? box.x : -box.y - box.height);
	const size = (box: BoundingBoxElement) => (horizontal ? box.width : box.height);
	const center = (box: BoundingBoxElement) => start(box) + size(box) / 2;
	const sorted = [...boxes].sort((left, right) => center(left) - center(right));
	const first = sorted[0]!,
		last = sorted.at(-1)!;
	const total = sorted.reduce((sum, box) => sum + size(box), 0);
	const gap = (start(last) + size(last) - start(first) - total) / (sorted.length - 1);
	let cursor = start(first);
	const positions = new Map<string, ElementPosition>();
	for (const box of sorted) {
		positions.set(box.id, horizontal ? { x: cursor } : { y: -cursor - box.height });
		cursor += size(box) + gap;
	}
	return positions;
}

/** The cached rectangular alignment box in page coordinates, including saved rotation/flips. */
function alignmentBox(shape: VisioShape): BoundingBoxElement | undefined {
	if (!shape.rotation || shape.kind !== 'shape' || !(shape.width > 0 && shape.height > 0))
		return undefined;
	const [a, b, c, d, e, f] = shape.transform;
	if (
		![a, b, c, d, e, f, shape.rotation.pinX, shape.rotation.pinY, shape.width, shape.height].every(
			Number.isFinite,
		)
	)
		return undefined;
	const points = [
		[0, 0],
		[shape.width, 0],
		[0, shape.height],
		[shape.width, shape.height],
	] as const;
	const xs = points.map(([x, y]) => a * x + c * y + e);
	const ys = points.map(([x, y]) => b * x + d * y + f);
	const x = Math.min(...xs),
		top = Math.max(...ys);
	const width = Math.max(...xs) - x,
		height = top - Math.min(...ys);
	if (![x, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0)
		return undefined;
	// Shared geometry uses downward-positive y, while Visio's page and pins use upward-positive y.
	return { id: shape.id, x, y: -top, width, height };
}

/**
 * Align to the first selected primary shape, or distribute equal edge gaps without guide glue.
 * https://learn.microsoft.com/en-us/office/vba/api/visio.selection.align
 * https://learn.microsoft.com/en-us/office/vba/api/visio.selection.distribute
 * Coarse candidates only: source locks, glue, formulas and dependencies remain authoritative.
 */
export function visioArrangeCommands(
	page: VisioPage,
	shapeIds: readonly string[],
	action: VisioArrangement,
): VisioEdit[] | undefined {
	const minimum = action.type === 'align' ? 2 : 3;
	if (
		shapeIds.length < minimum ||
		shapeIds.length > 1000 ||
		new Set(shapeIds).size !== shapeIds.length
	)
		return undefined;
	if (
		action.type === 'align'
			? !['left', 'center', 'right', 'top', 'middle', 'bottom'].includes(action.edge)
			: action.type !== 'distribute' || !['horizontal', 'vertical'].includes(action.axis)
	)
		return undefined;
	const shapes: VisioShape[] = [];
	const boxes: BoundingBoxElement[] = [];
	for (const id of shapeIds) {
		const shape = visioStyleFormattingShape(page, id);
		const box = shape && alignmentBox(shape);
		if (
			!shape ||
			!box ||
			page.connectors.some(
				(connection) => connection.fromShapeId === id || connection.toShapeId === id,
			)
		)
			return undefined;
		shapes.push(shape);
		boxes.push(box);
	}
	const reference = boxes[0]!;
	const normalized = boxes.map((box) => ({
		...box,
		x: box.x - reference.x,
		y: box.y - reference.y,
	}));
	const edge: AlignEdge | undefined =
		action.type === 'align' ? (action.edge === 'center' ? 'centerH' : action.edge) : undefined;
	const positions =
		action.type === 'align'
			? alignElements(normalized, edge!, {
					reference: 'slide',
					slideSize: { width: reference.width, height: reference.height },
				})
			: distributeBoxes(normalized, action.axis);
	const commands: VisioEdit[] = [];
	try {
		for (let index = 0; index < shapes.length; index++) {
			const shape = shapes[index]!,
				box = normalized[index]!;
			const position = positions.get(shape.id);
			if (!position) return undefined;
			const dx = (position.x ?? box.x) - box.x,
				dy = -((position.y ?? box.y) - box.y);
			// Preserve the primary anchor and distribution endpoints without rewriting protected pins.
			if (Math.abs(dx) < 1e-10 && Math.abs(dy) < 1e-10) continue;
			commands.push(
				visioPageEditToDrawing(page, {
					type: 'move-shape',
					pageId: page.id,
					shapeId: shape.id,
					x: shape.rotation!.pinX + dx,
					y: shape.rotation!.pinY + dy,
				}),
			);
		}
	} catch {
		return undefined;
	}
	return commands;
}
