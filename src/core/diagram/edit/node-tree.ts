// The content nodes of a SmartArt data model in tree order: each parent before its children, and
// siblings by their `parOf` connection's `srcOrd`. The editing operations and the data writer read
// array order as sibling order, so host models keep nodes in this order. Moved from
// `docx/diagram-nodes.ts` so Word and Excel read nodes the same way.
import type { DiagramDataModel } from '../types';
import type { DiagramHostNode } from './node-edit';

/** `nodes` reordered parent first, siblings in array order (a stable depth-first walk). */
export function treeOrderDiagramNodes<N extends { id: string; parentId?: string | undefined }>(
	nodes: readonly N[],
): N[] {
	const ids = new Set(nodes.map((node) => node.id));
	const childrenOf = new Map<string | undefined, N[]>();
	for (const node of nodes) {
		const parent = node.parentId && ids.has(node.parentId) ? node.parentId : undefined;
		const list = childrenOf.get(parent) ?? [];
		list.push(node);
		childrenOf.set(parent, list);
	}
	const ordered: N[] = [];
	const seen = new Set<string>();
	const visit = (parent: string | undefined): void => {
		for (const node of childrenOf.get(parent) ?? []) {
			if (seen.has(node.id)) continue;
			seen.add(node.id);
			ordered.push(node);
			visit(node.id);
		}
	};
	visit(undefined);
	// A cycle (malformed parents) leaves nodes unvisited: keep them, in their original order.
	for (const node of nodes) if (!seen.has(node.id)) ordered.push(node);
	return ordered;
}

/** The model's content nodes, with text, paragraphs and parents, in tree order. */
export function diagramNodesInTreeOrder(data: DiagramDataModel): DiagramHostNode[] {
	const order = new Map<string, number>();
	for (const connection of data.connections)
		if ((connection.type ?? 'parOf') === 'parOf' && !order.has(connection.destId))
			order.set(connection.destId, connection.srcOrd ?? 0);
	const nodes = data.nodes.map((node, index) => ({
		index,
		node: {
			id: node.modelId,
			text: node.text,
			...(node.textParagraphs ? { paragraphs: node.textParagraphs } : {}),
			...(data.parentById.get(node.modelId)
				? { parentId: data.parentById.get(node.modelId) as string }
				: {}),
			...(node.type === 'asst' ? { nodeType: 'asst' as const } : {}),
		} satisfies DiagramHostNode,
	}));
	nodes.sort(
		(a, b) =>
			(order.get(a.node.id) ?? Number.MAX_SAFE_INTEGER) -
				(order.get(b.node.id) ?? Number.MAX_SAFE_INTEGER) || a.index - b.index,
	);
	return treeOrderDiagramNodes(nodes.map((entry) => entry.node));
}
