import type { FormulaNode } from './graph.js';
import { stronglyConnected } from './scc.js';

/** A validated dependency order; formula results are still evaluated on every pass. */
export interface FormulaOrder {
	components: { nodes: FormulaNode[]; cyclic: boolean }[];
	ranks: Map<FormulaNode, number>;
}

export function formulaOrder(
	nodes: Iterable<FormulaNode>,
	successors: (node: FormulaNode) => FormulaNode[],
): FormulaOrder {
	const ranks = new Map<FormulaNode, number>();
	const components = stronglyConnected(nodes, successors).map((members, rank) => {
		for (const member of members) ranks.set(member, rank);
		const first = members[0] as FormulaNode;
		return { nodes: members, cyclic: members.length > 1 || successors(first).includes(first) };
	});
	return { components, ranks };
}
