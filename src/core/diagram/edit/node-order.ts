/**
 * SmartArt node reordering and level changes over the neutral data model:
 * move up/down, move to an index, promote and demote. Pure, like the other
 * node operations: each returns a new data object of the caller's type with
 * the drawing marked dirty and its shapes stale. Moved from
 * `pptx/core/utils/smartart-editing-node-ops.ts` and made generic.
 *
 * @module diagram/edit/node-order
 */

import { withEditedStructure } from './node-ops';
import type { DiagramEditableData } from './types';

/**
 * Move a node to a different position within its sibling group.
 * `direction` of 1 moves the node down/right, -1 moves it up/left.
 */
export function moveDiagramNode<D extends DiagramEditableData>(
	data: D,
	nodeId: string,
	direction: 1 | -1,
): D {
	const node = data.nodes.find((n) => n.id === nodeId);
	if (!node) {
		return data;
	}

	// Find siblings (nodes with the same parentId)
	const siblings = data.nodes.filter((n) => n.parentId === node.parentId);
	const currentIndex = siblings.findIndex((n) => n.id === nodeId);
	const targetIndex = currentIndex + direction;
	const current = siblings[currentIndex];
	const target = siblings[targetIndex];
	if (!current || !target) {
		return data;
	}

	// Swap in the full node list
	const nodes = [...data.nodes];
	const currentGlobalIndex = nodes.findIndex((n) => n.id === current.id);
	const targetGlobalIndex = nodes.findIndex((n) => n.id === target.id);
	if (currentGlobalIndex >= 0 && targetGlobalIndex >= 0) {
		nodes[currentGlobalIndex] = target;
		nodes[targetGlobalIndex] = current;
	}
	return withEditedStructure(data, nodes);
}

/**
 * Move a node to a specific index within its sibling group.
 *
 * Siblings are all nodes sharing the same `parentId`.
 * The node is removed from its current position among siblings and
 * re-inserted at `newIndex` (clamped to valid range).
 */
export function moveDiagramNodeToIndex<D extends DiagramEditableData>(
	data: D,
	nodeId: string,
	newIndex: number,
): D {
	const node = data.nodes.find((n) => n.id === nodeId);
	if (!node) {
		return data;
	}

	// Collect siblings in their original order
	const siblings = data.nodes.filter((n) => n.parentId === node.parentId);
	const currentIndex = siblings.findIndex((n) => n.id === nodeId);
	if (currentIndex < 0) {
		return data;
	}

	// Clamp the target index
	const clampedIndex = Math.max(0, Math.min(newIndex, siblings.length - 1));
	if (clampedIndex === currentIndex) {
		return data;
	}

	// Reorder siblings
	const reorderedSiblings = [...siblings];
	const [moved] = reorderedSiblings.splice(currentIndex, 1);
	if (moved) reorderedSiblings.splice(clampedIndex, 0, moved);

	// Rebuild the full node list preserving non-sibling positions
	const siblingIds = new Set(siblings.map((s) => s.id));
	const nodes: D['nodes'] = [];
	let sibIdx = 0;
	for (const n of data.nodes) {
		nodes.push(siblingIds.has(n.id) ? (reorderedSiblings[sibIdx++] ?? n) : n);
	}
	return withEditedStructure(data, nodes);
}

/** Promote a child node to be a sibling of its parent. */
export function promoteDiagramNode<D extends DiagramEditableData>(data: D, nodeId: string): D {
	const node = data.nodes.find((n) => n.id === nodeId);
	if (!node || !node.parentId) {
		return data;
	}

	const parent = data.nodes.find((n) => n.id === node.parentId);
	if (!parent) {
		return data;
	}

	const nodes = data.nodes.map((n) => (n.id === nodeId ? { ...n, parentId: parent.parentId } : n));

	// Update connections
	const connections = (data.connections ?? [])
		.filter((c) => !(c.sourceId === node.parentId && c.destId === nodeId))
		.map((c) => ({ ...c }));

	if (parent.parentId) {
		connections.push({
			sourceId: parent.parentId,
			destId: nodeId,
			type: 'parOf',
			srcOrd: 0,
			destOrd: 0,
		});
	}
	return withEditedStructure(data, nodes, connections);
}

/** Demote a node to become a child of its preceding sibling. */
export function demoteDiagramNode<D extends DiagramEditableData>(data: D, nodeId: string): D {
	const node = data.nodes.find((n) => n.id === nodeId);
	if (!node) {
		return data;
	}

	// Find the preceding sibling; the first sibling cannot be demoted.
	const siblings = data.nodes.filter((n) => n.parentId === node.parentId);
	const currentIndex = siblings.findIndex((n) => n.id === nodeId);
	const previous = currentIndex > 0 ? siblings[currentIndex - 1] : undefined;
	if (!previous) {
		return data;
	}

	const newParentId = previous.id;
	const nodes = data.nodes.map((n) => (n.id === nodeId ? { ...n, parentId: newParentId } : n));

	// Update connections
	const connections = (data.connections ?? [])
		.filter((c) => !(c.sourceId === node.parentId && c.destId === nodeId))
		.map((c) => ({ ...c }));

	connections.push({
		sourceId: newParentId,
		destId: nodeId,
		type: 'parOf',
		srcOrd: 0,
		destOrd: 0,
	});
	return withEditedStructure(data, nodes, connections);
}
