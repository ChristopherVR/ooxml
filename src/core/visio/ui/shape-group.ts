import type { VisioPage, VisioShape } from '../model';
import type { VisioGroupShapesEdit, VisioUngroupShapeEdit } from '../edit-group-commands';
import { visioLocalRotationShape } from './shape-rotation';
import { visioNextShapeId } from './shape-id';

/** Model-level scope of Group and Ungroup: visible, unlayered, local and unglued 2D trees.
 * Source locks, formulas and glue remain authoritative in core admission.
 */
function groupable(page: VisioPage, id: string): VisioShape | undefined {
	const shape = visioLocalRotationShape(page, id);
	return shape && !shape.layerIds?.length ? shape : undefined;
}

/** A top-level group the pointer can move as one sheet. */
export function visioMovableGroup(page: VisioPage, id: string): VisioShape | undefined {
	const shape = groupable(page, id);
	return shape?.kind === 'group' && shape.children.length ? shape : undefined;
}

/** Group two or more selected top-level shapes; members keep page stacking order. */
export function visioGroupCommand(
	page: VisioPage,
	shapeIds: readonly string[],
): VisioGroupShapesEdit | undefined {
	if (shapeIds.length < 2 || shapeIds.length > 1000 || new Set(shapeIds).size !== shapeIds.length)
		return undefined;
	const wanted = new Set(shapeIds);
	const members = page.shapes.filter((shape) => wanted.has(shape.id));
	if (members.length !== wanted.size || members.some((shape) => !groupable(page, shape.id)))
		return undefined;
	let shapeId: string;
	try {
		shapeId = visioNextShapeId(page);
	} catch {
		return undefined;
	}
	return {
		type: 'group-shapes',
		pageId: page.id,
		shapeId,
		memberIds: members.map((shape) => shape.id),
	};
}

/** Ungroup exactly one selected top-level group. */
export function visioUngroupCommand(
	page: VisioPage,
	shapeIds: readonly string[],
): VisioUngroupShapeEdit | undefined {
	if (shapeIds.length !== 1 || !visioMovableGroup(page, shapeIds[0]!)) return undefined;
	return { type: 'ungroup-shape', pageId: page.id, shapeId: shapeIds[0]! };
}
