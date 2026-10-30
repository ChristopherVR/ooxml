import type { VisioPage, VisioShape } from '../model';

const pageShape = (page: VisioPage, id: string): VisioShape | undefined => {
	const matches = page.shapes.filter((shape) => shape.id === id);
	return matches.length === 1 ? matches[0] : undefined;
};

/** Every shape ID inside a shape, itself included. */
function subtree(shape: VisioShape, ids = new Set<string>()): Set<string> {
	ids.add(shape.id);
	for (const child of shape.children) subtree(child, ids);
	return ids;
}

/**
 * The part of a selection that is moved or resized as shapes. A connector in the selection with
 * both ends glued to selected shapes (or to members of selected groups) is not moved itself: the
 * core lays it out again after each shape command, which puts it where its two ends now are. A
 * connector with a free end, or an end glued outside the selection, makes the whole selection
 * ineligible (`undefined`): its free end has no rule here, and Visio moves it differently.
 */
export function visioSelectionShapes(
	page: VisioPage,
	ids: readonly string[],
): string[] | undefined {
	const connectors = ids.filter((id) => pageShape(page, id)?.kind === 'connector');
	if (!connectors.length) return [...ids];
	const shapes = ids.filter((id) => !connectors.includes(id));
	if (!shapes.length) return undefined;
	const covered = new Set<string>();
	for (const id of shapes) {
		const shape = pageShape(page, id);
		if (!shape) return undefined;
		subtree(shape, covered);
	}
	for (const id of connectors) {
		const ends = new Set<string>();
		for (const connection of page.connectors) {
			if (connection.fromShapeId !== id) continue;
			if (!covered.has(connection.toShapeId)) return undefined;
			ends.add(connection.fromCell);
		}
		if (!ends.has('BeginX') || !ends.has('EndX')) return undefined;
	}
	return shapes;
}
