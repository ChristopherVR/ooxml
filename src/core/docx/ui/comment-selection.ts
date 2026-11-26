import type { Node } from 'prosemirror-model';
import type { Selection } from 'prosemirror-state';
import { fieldsBalanced } from './field-guard';
import { fieldResultRanges } from './field-results';

const words = new Intl.Segmenter(undefined, { granularity: 'word' });

/** Resolve a collapsed text selection to the word under the caret, or retain the point. */
export function commentTargetRange(doc: Node, selection: Selection): { from: number; to: number } {
	const { $from, from, to, empty } = selection;
	if (!empty) return commentSelectionRange(doc, from, to);
	if (!$from.parent.isTextblock) return { from, to };
	let text = '';
	$from.parent.forEach((node) => {
		text += node.text ?? '￼'.repeat(node.nodeSize);
	});
	const word = words.segment(text).containing($from.parentOffset);
	if (!word?.isWordLike) return { from, to };
	const start = $from.start() + word.index;
	return commentSelectionRange(doc, start, start + word.segment.length);
}

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
