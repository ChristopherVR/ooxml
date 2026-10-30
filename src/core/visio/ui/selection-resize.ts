import { applyResize } from '../../geometry/element-transform';
import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioPageEditToDrawing } from './page-edit';
import type { VisioPagePoint } from './shape-move';
import {
	VISIO_RESIZE_HANDLES,
	visioResizeDrag,
	visioResizeFrame,
	visioResizeShape,
	visioProportionalSize,
	visioResizeKeepsProportions,
	type VisioResizeDragOptions,
	type VisioResizeFrame,
	type VisioResizeHandle,
} from './shape-resize';

function candidates(page: VisioPage, ids: readonly string[]): VisioShape[] | undefined {
	if (!ids.length || ids.length > 500 || new Set(ids).size !== ids.length) return undefined;
	const shapes: VisioShape[] = [];
	for (const id of ids) {
		const shape = visioResizeShape(page, id);
		if (!shape) return undefined;
		// Anisotropic page scaling would shear rotated shapes. Local groups, stencil instances
		// and attached connectors need separate selection-resize proofs, even if singly resizable.
		if (ids.length > 1) {
			const [a, b, c, d] = shape.transform;
			if (
				shape.children.length ||
				shape.masterId ||
				shape.kind !== 'shape' ||
				Math.abs(b) > 1e-10 ||
				Math.abs(c) > 1e-10 ||
				Math.abs(Math.abs(a) - 1) > 1e-10 ||
				Math.abs(Math.abs(d) - 1) > 1e-10 ||
				page.connectors.some((edge) => edge.fromShapeId === id || edge.toShapeId === id)
			)
				return undefined;
		}
		shapes.push(shape);
	}
	return shapes;
}

function selectionFrame(shapes: readonly VisioShape[]): VisioResizeFrame | undefined {
	const first = shapes[0]!;
	if (shapes.length === 1) return visioResizeFrame(first, first, { x: 0, y: 0 });
	let left = Infinity,
		bottom = Infinity,
		right = -Infinity,
		top = -Infinity;
	for (const shape of shapes) {
		const [a, , , d, x, y] = shape.transform;
		left = Math.min(left, x, x + a * shape.width);
		right = Math.max(right, x, x + a * shape.width);
		bottom = Math.min(bottom, y, y + d * shape.height);
		top = Math.max(top, y, y + d * shape.height);
	}
	const width = right - left,
		height = top - bottom;
	return width > 0 && height > 0 && [left, bottom, width, height].every(Number.isFinite)
		? {
				width,
				height,
				transform: [1, 0, 0, 1, left, bottom],
				pinX: left + width / 2,
				pinY: bottom + height / 2,
			}
		: undefined;
}

/** One local frame for a single shape, or a page-aligned frame for eligible local leaves. */
export function visioResizeSelectionFrame(
	page: VisioPage,
	ids: readonly string[],
): VisioResizeFrame | undefined {
	const shapes = candidates(page, ids);
	return shapes && selectionFrame(shapes);
}

/** Scale members and spacing together. Commit all commands atomically: source locks, formulas
 * and dependencies remain authoritative. Before accepting the edited document, call
 * assertVisioGeometryExpectation with these members; discard the entire batch on rejection.
 */
export function visioResizeSelectionDrag(
	page: VisioPage,
	ids: readonly string[],
	handle: VisioResizeHandle,
	start: VisioPagePoint,
	end: VisioPagePoint,
	options: VisioResizeDragOptions = {},
):
	| {
			frame: VisioResizeFrame;
			commands: VisioEdit[];
			members: { id: string; frame: VisioResizeFrame }[];
	  }
	| undefined {
	const shapes = candidates(page, ids);
	if (!shapes || !VISIO_RESIZE_HANDLES.some((item) => item.id === handle)) return undefined;
	if (shapes.length === 1) {
		const result = visioResizeDrag(page, ids[0]!, handle, start, end, options);
		return (
			result && {
				frame: result.frame,
				commands: [result.command],
				members: [{ id: ids[0]!, frame: result.frame }],
			}
		);
	}
	const original = selectionFrame(shapes);
	if (!original || ![start.x, start.y, end.x, end.y].every(Number.isFinite)) return undefined;
	const [, , , , left, bottom] = original.transform;
	// Shared PowerPoint interaction geometry uses downward-positive y. Preserve the per-member
	// minimum (including already tiny shapes) by constraining each common scale independently.
	const min = {
		width:
			original.width *
			Math.max(...shapes.map((shape) => Math.min(1 / 16, shape.width) / shape.width)),
		height:
			original.height *
			Math.max(...shapes.map((shape) => Math.min(1 / 16, shape.height) / shape.height)),
	};
	// Ctrl moves both edges of the surrounding rectangle, so the drag counts twice.
	const reach = options.centre ? 2 : 1;
	let box = applyResize(
		{ x: left, y: -bottom - original.height, width: original.width, height: original.height },
		handle,
		reach * (end.x - start.x),
		reach * (end.y - start.y),
		1,
		{ minWidth: min.width, minHeight: min.height },
	);
	// One common scale: any member's LockAspect keeps the whole selection proportional.
	const aspect = shapes.some((shape) => shape.resizeLocks?.aspect);
	if (visioResizeKeepsProportions(handle, { ...options, aspect })) {
		// Scale the surrounding rectangle uniformly about the point opposite the handle (or its
		// centre with Ctrl).
		const location = VISIO_RESIZE_HANDLES.find((item) => item.id === handle)!;
		const size = visioProportionalSize(original, box, location, min);
		const ax = options.centre ? 0.5 : 1 - location.x,
			ay = options.centre ? 0.5 : 1 - location.y;
		const nextLeft = left + ax * (original.width - size.width),
			nextBottomEdge = bottom + ay * (original.height - size.height);
		box = { ...box, ...size, x: nextLeft, y: -nextBottomEdge - size.height };
	}
	if (options.centre) {
		// Ctrl keeps the surrounding rectangle's centre fixed (box y holds the negated top edge).
		box = {
			...box,
			x: left + (original.width - box.width) / 2,
			y: -(bottom + original.height / 2) - box.height / 2,
		};
	}
	const sx = box.width / original.width,
		sy = box.height / original.height;
	// A member's LockWidth or LockHeight refuses a common scale that would change that dimension.
	if (
		shapes.some(
			(shape) =>
				(shape.resizeLocks?.width && Math.abs(sx - 1) > 1e-12) ||
				(shape.resizeLocks?.height && Math.abs(sy - 1) > 1e-12),
		)
	)
		return undefined;
	const nextBottom = -box.y - box.height;
	const frame: VisioResizeFrame = {
		width: box.width,
		height: box.height,
		transform: [1, 0, 0, 1, box.x, nextBottom],
		pinX: box.x + box.width / 2,
		pinY: nextBottom + box.height / 2,
	};
	if (
		![...frame.transform, frame.width, frame.height, frame.pinX, frame.pinY].every(Number.isFinite)
	)
		return undefined;
	const commands: VisioEdit[] = [];
	const members: { id: string; frame: VisioResizeFrame }[] = [];
	try {
		for (const shape of shapes) {
			const width = shape.width * sx,
				height = shape.height * sy;
			const x = box.x + (shape.rotation!.pinX - left) * sx;
			const y = nextBottom + (shape.rotation!.pinY - bottom) * sy;
			// Use the lower-left local anchor, whose resulting pin is proved by the same helper
			// as the source writer, then move to the affine-mapped pin. This also handles off-center
			// local pins. Skip a redundant move if that anchor already lands here.
			const resized = visioResizeFrame(shape, { width, height }, { x: 0, y: 0 });
			if (!resized) return undefined;
			members.push({
				id: shape.id,
				frame: {
					...resized,
					pinX: x,
					pinY: y,
					transform: [
						resized.transform[0],
						resized.transform[1],
						resized.transform[2],
						resized.transform[3],
						resized.transform[4] + x - resized.pinX,
						resized.transform[5] + y - resized.pinY,
					],
				},
			});
			if (Math.abs(width - shape.width) > 1e-10 || Math.abs(height - shape.height) > 1e-10) {
				const command = visioPageEditToDrawing(page, {
					type: 'resize-shape',
					pageId: page.id,
					shapeId: shape.id,
					width,
					height,
					anchor: { x: 0, y: 0 },
				});
				if (command.type !== 'resize-shape' || command.width > 1e6 || command.height > 1e6)
					return undefined;
				commands.push(command);
			}
			if (Math.abs(x - resized.pinX) > 1e-10 || Math.abs(y - resized.pinY) > 1e-10) {
				const command = visioPageEditToDrawing(page, {
					type: 'move-shape',
					pageId: page.id,
					shapeId: shape.id,
					x,
					y,
				});
				if (command.type !== 'move-shape' || ![command.x, command.y].every(Number.isFinite))
					return undefined;
				commands.push(command);
			}
		}
	} catch {
		return undefined;
	}
	return { frame, commands, members };
}
