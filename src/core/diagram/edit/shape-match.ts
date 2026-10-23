/**
 * Drawing-shape bookkeeping shared by the SmartArt node operations: which
 * shape shows a node, and how an edit marks the shapes stale. Moved from
 * `pptx/core/utils/smartart-editing-node-ops.ts` and
 * `smartart-editing-node-style.ts`, which each kept a copy of the matcher.
 *
 * @module diagram/edit/shape-match
 */

import type { DiagramNode } from '../model/node';
import type { DiagramEditShape } from './types';

/**
 * Mark drawing shapes as stale after a structural node edit.
 *
 * - If the element previously had drawing shapes (from file load or a prior
 *   reflow), returns an empty array `[]` to signal that a rebuild is needed.
 * - If the element never had drawing shapes (`undefined`), returns `undefined`
 *   so the family renderer continues handling display without triggering
 *   a lossy reflow (e.g. polygon-to-chevron downgrade for pyramid layouts).
 */
export function markDiagramShapesStale<S>(shapes: S[] | undefined): S[] | undefined {
	return shapes !== undefined && shapes.length > 0 ? [] : undefined;
}

/**
 * Find the drawing shape index that corresponds to a given node ID: a reflow
 * shape's id suffix, else a 1:1 positional mapping when the counts align, else
 * a unique non-empty text match. `-1` when none resolves (PowerPoint-generated
 * shapes with opaque ids).
 */
export function findDiagramShapeIndex(
	shapes: readonly DiagramEditShape[],
	nodes: readonly DiagramNode[],
	nodeId: string,
): number {
	// 1. Reflow shapes embed the node id as the id suffix.
	for (let i = 0; i < shapes.length; i++) {
		const id = shapes[i]!.id;
		if (id.startsWith('reflow-') && id.endsWith(`-${nodeId}`)) {
			return i;
		}
	}
	// 2. 1:1 positional mapping when counts align.
	if (shapes.length === nodes.length) {
		const nodeIdx = nodes.findIndex((n) => n.id === nodeId);
		if (nodeIdx >= 0) {
			return nodeIdx;
		}
	}
	// 3. Unique non-empty text match.
	const node = nodes.find((n) => n.id === nodeId);
	if (node?.text) {
		const text = node.text.trim();
		const matches: number[] = [];
		for (let i = 0; i < shapes.length; i++) {
			if (shapes[i]!.text?.trim() === text) {
				matches.push(i);
			}
		}
		if (matches.length === 1) {
			return matches[0]!;
		}
	}
	return -1;
}

/**
 * Replace the shape that shows `nodeId` with `patch(shape)` (immutable copy).
 * Shapes that are absent or empty are returned unchanged; an unresolvable
 * match returns `[]` so a full rebuild picks the change up.
 */
export function patchDiagramShape<S extends DiagramEditShape>(
	shapes: S[] | undefined,
	nodes: readonly DiagramNode[],
	nodeId: string,
	patch: (shape: S) => S,
): S[] | undefined {
	if (!shapes || shapes.length === 0) {
		return shapes;
	}
	const idx = findDiagramShapeIndex(shapes, nodes, nodeId);
	if (idx < 0) {
		return [];
	}
	const updated = [...shapes];
	updated[idx] = patch(updated[idx]!);
	return updated;
}
