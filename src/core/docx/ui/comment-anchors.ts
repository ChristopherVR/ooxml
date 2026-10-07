import type { Mark, Node } from 'prosemirror-model';
import { inlineRunProperties } from './inline-run-properties';

/** Union comment anchors from independent marks, including legacy grouped-id marks. */
export function commentIdsFromMarks(marks: readonly Mark[]): string[] {
	const ids = new Set<string>();
	for (const mark of marks) {
		if (mark.type.name !== 'comment' || !Array.isArray(mark.attrs.ids)) continue;
		for (const id of mark.attrs.ids) if (typeof id === 'string' && id) ids.add(id);
	}
	return [...ids].sort();
}

/** Imported inline elements retain anchors in run metadata rather than element marks. */
export function commentIdsFromNode(node: Node): string[] {
	const ids = new Set(commentIdsFromMarks(node.marks));
	const imported =
		!node.isText && node.type.spec.attrs?.format ? inlineRunProperties(node).commentIds : undefined;
	for (const id of imported ?? []) if (typeof id === 'string' && id) ids.add(id);
	return [...ids].sort();
}
