import type { VisioEdit } from '../edit';
import type { VisioPage, VisioShape } from '../model';
/** The model-level scope of local rotation. Source protections remain enforced by core edits. */
export function visioLocalRotationShape(page: VisioPage, shapeId: string): VisioShape | undefined {
	const shape = page.shapes.find((shape) => shape.id === shapeId);
	if (
		!shape ||
		shape.kind !== 'shape' ||
		shape.children.length ||
		shape.masterId ||
		shape.hidden ||
		!shape.rotation ||
		!(shape.width > 0 && shape.height > 0) ||
		page.connectors.some(
			(connection) => connection.fromShapeId === shape.id || connection.toShapeId === shape.id,
		)
	)
		return undefined;
	return shape;
}
/** Native per-shape quarter turns normalize to signed angles while retaining the saved pin. */
export function visioQuarterTurnCommand(
	page: VisioPage,
	shapeId: string,
	direction: 'left' | 'right',
): Extract<VisioEdit, { type: 'rotate-shape' }> | undefined {
	const shape = visioLocalRotationShape(page, shapeId);
	if (!shape?.rotation || !Number.isFinite(shape.rotation.angle)) return undefined;
	const angle = shape.rotation.angle + ((direction === 'left' ? 1 : -1) * Math.PI) / 2;
	let normalized = Math.atan2(Math.sin(angle), Math.cos(angle));
	// Native XML rounds angle caches; keep the signed half-turn rather than crossing the branch cut.
	if (Math.abs(Math.abs(normalized) - Math.PI) < 1e-12) normalized = Math.sign(angle) * Math.PI;
	return {
		type: 'rotate-shape',
		pageId: page.id,
		shapeId,
		angle: normalized,
	};
}
