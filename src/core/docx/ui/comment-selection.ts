import type { Node } from 'prosemirror-model';
import { fieldsBalanced } from './field-guard';
import { fieldResultRanges } from './field-results';

/** Native Word anchors a selected field marker or result to the complete field. */
export function commentSelectionRange(
	doc: Node,
	from: number,
	to: number,
): { from: number; to: number } {
	if (from === to) return { from, to };
	const stack: number[] = [];
	const fields: { from: number; to: number }[] = fieldResultRanges(doc).filter(
		(range) => range.mark.attrs.simple,
	);
	if (fieldsBalanced(doc))
		doc.descendants((node, pos) => {
			if (node.type.name !== 'fieldMarker') return;
			if (node.attrs.kind === 'begin') stack.push(pos);
			else if (node.attrs.kind === 'end') {
				const start = stack.pop();
				if (start !== undefined) fields.push({ from: start, to: pos + node.nodeSize });
			}
		});
	for (const field of fields) {
		if (from < field.to && to > field.from) {
			from = Math.min(from, field.from);
			to = Math.max(to, field.to);
		}
	}
	return { from, to };
}
