import type { VisioPage, VisioShape } from '../model';

/** Coarse scene admission; source display bands and protection are checked by editVsdx. */
export function visioOrderingShape(page: VisioPage, shapeId: string): VisioShape | undefined {
	if (
		page.shapes.some(
			(shape) =>
				shape.kind !== 'shape' ||
				shape.masterId ||
				shape.children.length ||
				shape.image ||
				shape.foreignVector,
		)
	)
		return undefined;
	const shape = page.shapes.find((candidate) => candidate.id === shapeId);
	return shape && !shape.hidden && !shape.layerIds?.length ? shape : undefined;
}
