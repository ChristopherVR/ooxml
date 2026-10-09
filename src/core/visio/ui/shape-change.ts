import type { VisioPage, VisioShape } from '../model';
import type { VisioChangeShapeEdit, VisioChangeShapeTarget } from '../edit-change-shape-commands';

/**
 * Coarse scene admission for Home > Change Shape: the reason the selection cannot change shape,
 * or undefined when it can be tried. Locks, formula dependencies and control handles are checked
 * by `editVsdx`, which refuses the edit with its own reason.
 */
export function visioChangeShapeRefusal(
	page: VisioPage | undefined,
	selection: readonly Pick<VisioShape, 'id'>[],
): string | undefined {
	if (!page) return 'Open a drawing page.';
	if (selection.length !== 1) return 'Select exactly one shape.';
	const shape = page.shapes.find((candidate) => candidate.id === selection[0]!.id);
	if (!shape) return 'Select a top-level shape on this page; shapes inside groups cannot change.';
	if (shape.hidden) return 'The selected shape is hidden.';
	if (shape.masterId)
		return 'Master shapes inherit their geometry; replacing the master is not supported yet.';
	if (shape.kind === 'group' || shape.children.length)
		return 'Groups cannot change shape; change their member shapes instead.';
	if (shape.kind === 'connector') return 'Lines and connectors (1D shapes) cannot change shape.';
	if (shape.kind === 'foreign' || shape.image || shape.foreignVector)
		return 'Pictures and embedded objects have no outline to change.';
	if (!(shape.width > 0 && shape.height > 0)) return 'The shape has no width or height.';
	if (
		page.connectors.some(
			(connection) => connection.fromShapeId === shape.id || connection.toShapeId === shape.id,
		)
	)
		return 'Glued connectors would need rerouting, which is not supported yet.';
	return undefined;
}

/** The change-shape command for the single selected shape, when the scene admits it. */
export function visioChangeShapeCommand(
	page: VisioPage | undefined,
	selection: readonly Pick<VisioShape, 'id'>[],
	target: VisioChangeShapeTarget,
): VisioChangeShapeEdit | undefined {
	if (!page || visioChangeShapeRefusal(page, selection) !== undefined) return undefined;
	return { type: 'change-shape', pageId: page.id, shapeId: selection[0]!.id, shape: target };
}
