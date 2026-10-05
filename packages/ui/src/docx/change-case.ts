import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { WORD, transformCase, type CaseMode } from 'ooxml-core/docx';
import { schema } from './schema';

export type { CaseMode } from 'ooxml-core/docx';

/** The selection, or with a caret the word around it (as Word does). Null when there is none. */
function targetRange(view: EditorView): { from: number; to: number } | null {
	const { from, to, empty, $from } = view.state.selection;
	if (!empty) return { from, to };
	const parent = $from.parent;
	if (!parent.isTextblock) return null;
	const text = parent.textBetween(0, parent.content.size, undefined, '￼');
	let start = $from.parentOffset;
	let end = start;
	while (start > 0 && WORD.test(text[start - 1] ?? '')) start--;
	while (end < text.length && WORD.test(text[end] ?? '')) end++;
	if (start === end) return null;
	const base = $from.start();
	return { from: base + start, to: base + end };
}

/** Rewrites the case of the selected text, keeping each run's marks. Returns false when nothing changed. */
export function changeCase(view: EditorView, mode: CaseMode): boolean {
	if (!view.editable) return false;
	const range = targetRange(view);
	if (!range) return false;
	const edits: Array<{
		from: number;
		to: number;
		text: string;
		marks: readonly import('prosemirror-model').Mark[];
	}> = [];
	let sentenceStart = true;
	view.state.doc.nodesBetween(range.from, range.to, (node, pos) => {
		if (node.isTextblock) sentenceStart = true;
		if (!node.isText) return;
		const from = Math.max(pos, range.from);
		const to = Math.min(pos + node.nodeSize, range.to);
		const slice = node.text!.slice(from - pos, to - pos);
		const text = transformCase(slice, mode, sentenceStart);
		sentenceStart = /[.!?]\s*$/.test(slice);
		if (text !== slice) edits.push({ from, to, text, marks: node.marks });
	});
	if (!edits.length) return false;
	let tr = view.state.tr;
	for (const edit of edits.reverse())
		tr = tr.replaceWith(edit.from, edit.to, schema.text(edit.text, edit.marks));
	const { from, to, empty } = view.state.selection;
	const caret = Math.min(from, tr.doc.content.size);
	view.dispatch(
		tr.setSelection(
			TextSelection.create(
				tr.doc,
				empty ? caret : tr.mapping.map(from, -1),
				empty ? caret : tr.mapping.map(to, 1),
			),
		),
	);
	return true;
}
