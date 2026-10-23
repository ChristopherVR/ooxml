/**
 * The text pane's view of a diagram: each node's depth below the top level. A top-level node is
 * one without a parent or whose parent (the document point) is not in the list, the shape the
 * loaders produce. Parents precede their children in document order.
 *
 * @module diagram/edit/outline
 */
import type { DiagramNode } from '../model/node';

/** Each node's depth, in list order: 0 for a top-level node, its parent's depth + 1 otherwise. */
export function diagramNodeLevels(
	nodes: readonly Pick<DiagramNode, 'id' | 'parentId'>[],
): number[] {
	const ids = new Set(nodes.map((node) => node.id));
	const levels = new Map<string, number>();
	return nodes.map((node) => {
		const parent = node.parentId && ids.has(node.parentId) ? node.parentId : undefined;
		const level = parent === undefined ? 0 : (levels.get(parent) ?? 0) + 1;
		levels.set(node.id, level);
		return level;
	});
}
