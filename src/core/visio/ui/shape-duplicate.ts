import type { VisioPage } from '../model';
import type { VisioDuplicateShapesEdit } from '../edit-duplicate-commands';
import { visioNextShapeId } from './shape-id';
import { visioLocalFormattingShape } from './formatting';

/** Allocate against the complete page tree; source guards remain authoritative. */
export function visioDuplicateCommand(
	page: VisioPage,
	shapeIds: readonly string[],
): VisioDuplicateShapesEdit | undefined {
	if (
		!shapeIds.length ||
		shapeIds.length > 1000 ||
		new Set(shapeIds).size !== shapeIds.length ||
		shapeIds.some((id) => {
			const shape = visioLocalFormattingShape(page, id);
			return (
				!shape ||
				shape.kind !== 'shape' ||
				!(shape.width > 0) ||
				!(shape.height > 0) ||
				page.connectors.some(
					(connection) => connection.fromShapeId === id || connection.toShapeId === id,
				)
			);
		})
	)
		return undefined;
	let next: number;
	try {
		next = Number(visioNextShapeId(page));
	} catch {
		return undefined;
	}
	if (next + shapeIds.length - 1 > 0xffffffff) return undefined;
	const selected = new Set(shapeIds);
	const mapping = new Map(
		page.shapes
			.filter((shape) => selected.has(shape.id))
			.map((shape) => [shape.id, String(next++)]),
	);
	return {
		type: 'duplicate-shapes',
		pageId: page.id,
		copies: shapeIds.map((shapeId) => ({ shapeId, newShapeId: mapping.get(shapeId)! })),
		offsetX: 0.33,
		offsetY: -0.33,
	};
}
