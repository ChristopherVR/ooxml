import type { VisioEdit } from '../edit';
import type { VisioPage, VisioShape } from '../model';
/** The model-level scope of local rotation. Source protections remain enforced by core edits. */
export function visioLocalRotationShape(
	page: VisioPage,
	shapeId: string,
	includeGroups = true,
): VisioShape | undefined {
	const shape = page.shapes.find((shape) => shape.id === shapeId);
	if (!shape || shape.hidden) return undefined;
	const pending = [shape],
		ids = new Set<string>();
	while (pending.length) {
		const node = pending.pop()!;
		if (
			ids.size >= 10000 ||
			ids.has(node.id) ||
			node.masterId ||
			!node.rotation ||
			!(node.width > 0 && node.height > 0) ||
			![...node.transform, node.rotation.pinX, node.rotation.pinY, node.rotation.angle].every(
				Number.isFinite,
			) ||
			(node.kind === 'group'
				? !includeGroups || !node.children.length
				: node.kind !== 'shape' || !!node.children.length)
		)
			return undefined;
		ids.add(node.id);
		pending.push(...node.children);
	}
	if (
		page.connectors.some(
			(connection) =>
				ids.has(connection.fromShapeId) ||
				(ids.has(connection.toShapeId) && connection.toShapeId !== shapeId),
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
