import type { Node } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
/** Enforce the serializer's existing structure boundary before unsupported edits enter history. */
export function tableStructureGuardPlugin(): Plugin {
	const cache = new WeakMap<Node, ReadonlyMap<string, Node>>();
	const tables = (node: Node): ReadonlyMap<string, Node> => {
		const prior = cache.get(node);
		if (prior) return prior;
		const result = new Map<string, Node>();
		if (node.type.name === 'table') result.set(String(node.attrs.id), node);
		else if (!node.isTextblock && !node.isLeaf)
			node.forEach((child) => {
				for (const [id, table] of tables(child)) result.set(id, table);
			});
		cache.set(node, result);
		return result;
	};
	const sameStructure = (before: Node, after: Node): boolean => {
		if (before.childCount !== after.childCount) return false;
		for (let rowIndex = 0; rowIndex < before.childCount; rowIndex++) {
			const row = before.child(rowIndex),
				next = after.child(rowIndex);
			if (row.childCount !== next.childCount) return false;
			for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
				const cell = row.child(cellIndex),
					target = next.child(cellIndex);
				if (cell.type !== target.type || cell.childCount !== target.childCount) return false;
				for (const attr of ['sourceCellKey', 'colspan', 'rowspan', 'widthTwips'])
					if (cell.attrs[attr] !== target.attrs[attr]) return false;
				for (let i = 0; i < cell.childCount; i++) {
					const child = cell.child(i),
						other = target.child(i);
					if (child.type !== other.type) return false;
					if (child.type.name === 'paragraph') {
						if (child.attrs.id !== other.attrs.id) return false;
					} else if (child.type.name === 'table') {
						// A nested table keeps its grid and merges; only the text in its cells may change.
						if (!child.sameMarkup(other) || !sameStructure(child, other)) return false;
					} else if (!child.eq(other)) return false;
				}
			}
		}
		return true;
	};
	return new Plugin({
		filterTransaction(transaction, state) {
			if (!transaction.docChanged) return true;
			const after = tables(transaction.doc);
			for (const [id, before] of tables(state.doc)) {
				if (before.attrs.structureEditable !== false) continue;
				const target = after.get(id);
				// Deleting a complete source table is supported; changing its internal source slots is not.
				if (
					target &&
					target !== before &&
					(target.attrs.structureEditable !== false || !sameStructure(before, target))
				)
					return false;
			}
			return true;
		},
	});
}
