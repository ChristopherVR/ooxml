/**
 * The SmartArt node forest the hierarchy arrangers walk: nodes linked into
 * trees either from `parentId` pointers (the flat point list a reader
 * produces) or from already-nested `children` arrays, and the width and
 * depth measures the tree placers size rows with.
 *
 * Extracted from `pptx/core/utils/smartart-helpers.ts`; the pptx module
 * re-exports these next to its slide-element factories.
 */

import type { DiagramNode } from '../model';

/** One node of the forest and its child subtrees. */
export interface TreeNode<S = unknown> {
	node: DiagramNode<S>;
	children: TreeNode<S>[];
}

/** Link a flat node list into trees through `parentId` (orphans become roots). */
export function buildForest<S>(nodes: DiagramNode<S>[]): TreeNode<S>[] {
	const map = new Map<string, TreeNode<S>>();
	for (const n of nodes) {
		map.set(n.id, { node: n, children: [] });
	}
	const roots: TreeNode<S>[] = [];
	for (const n of nodes) {
		const tn = map.get(n.id)!;
		if (n.parentId && map.has(n.parentId)) {
			map.get(n.parentId)!.children.push(tn);
		} else {
			roots.push(tn);
		}
	}
	return roots;
}

/**
 * Build a forest from a node array, accepting either input shape:
 * - A flat list with `parentId` pointers (the format the parser emits for
 *   flat `<dgm:pt>` elements) - handled by delegating to {@link buildForest}.
 * - An already-nested list where every root node carries a `children` array.
 *
 * The DiagramML interpreter's hierarchy arranger receives either shape
 * depending on the caller, so it needs both branches; `buildForest` alone
 * only handles the flat case.
 */
export function buildTree<S>(nodes: DiagramNode<S>[]): TreeNode<S>[] {
	const hasNestedChildren = nodes.some((n) => n.children !== undefined && n.children.length > 0);
	if (!hasNestedChildren) {
		return buildForest(nodes);
	}
	const toTreeNode = (n: DiagramNode<S>): TreeNode<S> => ({
		node: n,
		children: (n.children ?? []).map(toTreeNode),
	});
	const allIds = new Set(nodes.map((n) => n.id));
	const roots = nodes.filter((n) => !n.parentId || !allIds.has(n.parentId));
	return roots.map(toTreeNode);
}

/**
 * Maximum recursion depth for tree traversal helpers.
 *
 * SmartArt diagrams are user-authored and may contain pathological depths
 * (or be the result of malformed input). Capping protects against stack
 * overflow while still accommodating any realistic SmartArt hierarchy.
 */
const MAX_TREE_DEPTH = 256;

/** Leaf count of `t` (a leaf counts as one). */
export function treeWidth<S>(t: TreeNode<S>, depth = 0): number {
	if (depth >= MAX_TREE_DEPTH) {
		return 1;
	}
	if (t.children.length === 0) {
		return 1;
	}
	let sum = 0;
	for (const c of t.children) {
		sum += treeWidth(c, depth + 1);
	}
	return sum;
}

/** Generation count of `t` (a leaf counts as one). */
export function treeDepth<S>(t: TreeNode<S>, depth = 0): number {
	if (depth >= MAX_TREE_DEPTH) {
		return 0;
	}
	if (t.children.length === 0) {
		return 1;
	}
	let max = 0;
	for (const c of t.children) {
		const d = treeDepth(c, depth + 1);
		if (d > max) {
			max = d;
		}
	}
	return 1 + max;
}
