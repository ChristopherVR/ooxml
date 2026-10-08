import { fail } from './package-common';

/** Append source-preserving local copies. Offsets use drawing inches, bottom-left origin. */
export interface VisioDuplicateShapesEdit {
	type: 'duplicate-shapes';
	pageId: string;
	/** Preserve selection order here; insertion follows original sibling stacking order. */
	copies: readonly { shapeId: string; newShapeId: string }[];
	offsetX: number;
	offsetY: number;
}
export function snapshotDuplicateShapes(edit: VisioDuplicateShapesEdit): VisioDuplicateShapesEdit {
	if (!Array.isArray(edit.copies) || !edit.copies.length || edit.copies.length > 1000)
		fail('INVALID_EDIT', 'Duplication requires between one and 1000 shape mappings.');
	if (
		[edit.offsetX, edit.offsetY].some(
			(value) => typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1e6,
		)
	)
		fail('INVALID_EDIT', 'Duplication offsets require bounded finite drawing inches.');
	const sources = new Set<string>(),
		targets = new Set<string>();
	const id = (value: unknown): string => {
		if (
			typeof value !== 'string' ||
			!/^(0|[1-9]\d{0,9})$/.test(value) ||
			Number(value) > 0xffffffff
		)
			fail('INVALID_EDIT', 'Duplicate shape IDs require canonical unsigned integers.');
		return value;
	};
	const copies = edit.copies.map((copy) => {
		if (!copy || typeof copy !== 'object') fail('INVALID_EDIT', 'Invalid duplicate shape mapping.');
		const shapeId = id(copy.shapeId),
			newShapeId = id(copy.newShapeId);
		if (sources.has(shapeId) || targets.has(newShapeId))
			fail('INVALID_EDIT', 'Duplicate shape mappings must have unique sources and targets.');
		sources.add(shapeId);
		targets.add(newShapeId);
		return { shapeId, newShapeId };
	});
	return {
		type: edit.type,
		pageId: edit.pageId,
		copies,
		offsetX: edit.offsetX,
		offsetY: edit.offsetY,
	};
}
