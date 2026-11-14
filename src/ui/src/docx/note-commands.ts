import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	formatNoteNumber,
	planNoteConversion,
	retargetNoteReference,
	type DocumentModel,
	type Note,
	type NoteConversionPlan,
	type NoteSelection,
	type Paragraph,
	type TextRun,
} from 'ooxml-core/docx';
import { schema } from './schema';
import {
	inlineNodeRun,
	notePartsJson,
	NotePartsStep,
	restoreNoteParts,
	runToInlineNodes,
} from 'ooxml-core/docx/ui';

export type NoteKind = 'footnote' | 'endnote';

const STYLES: Record<NoteKind, { text: string; reference: string }> = {
	footnote: { text: 'FootnoteText', reference: 'FootnoteReference' },
	endnote: { text: 'EndnoteText', reference: 'EndnoteReference' },
};

/** The next free note id: Word numbers user notes from 1 (separators use -1 and 0). */
export function nextNoteId(notes: Note[] | undefined): string {
	const used = (notes ?? []).map((note) => Number(note.id)).filter(Number.isFinite);
	return String(Math.max(0, ...used) + 1);
}

/** The formatting Word gives a note reference and a note's own number mark. */
function referenceFormatting(
	model: DocumentModel,
	kind: NoteKind,
): Pick<TextRun, 'verticalAlign' | 'style'> {
	const style = STYLES[kind].reference;
	return {
		verticalAlign: 'superscript',
		...(model.characterStyles?.styles[style] ? { style } : {}),
	};
}

/**
 * Word's Insert Footnote/Endnote: adds a note (number mark plus a space, in the note text style
 * when the document defines one) and inserts its superscript reference at the selection.
 * Returns the updated model and the new note's id.
 */
export function insertNote(
	view: EditorView,
	model: DocumentModel,
	kind: NoteKind,
): { model: DocumentModel; id: string } {
	const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
	const id = nextNoteId(model[key]);
	const formatting = referenceFormatting(model, kind);
	const textStyle = STYLES[kind].text;
	const paragraph: Paragraph = {
		type: 'paragraph',
		id: `${kind === 'footnote' ? 'fn' : 'en'}${id}-p0`,
		runs: [{ text: '', noteMark: kind, ...formatting }, { text: ' ' }],
		...(model.paragraphStyles?.styles[textStyle] ? { style: textStyle } : {}),
	};
	const next: DocumentModel = {
		...model,
		[key]: [...(model[key] ?? []), { id, blocks: [paragraph] }],
	};
	const reference = schema.nodes.noteReference.create({
		kind,
		id,
		number: 1,
		format: JSON.stringify(formatting),
	});
	view.dispatch(
		view.state.tr
			.replaceSelectionWith(reference, false)
			.setDocAttribute('noteParts', notePartsJson(next))
			.scrollIntoView(),
	);
	return { model: next, id };
}

/** The note reference touching the caret (a right-click beside it), or undefined. */
export function noteReferenceAtSelection(
	view: EditorView,
): { kind: NoteKind; id: string } | undefined {
	const { $from, empty } = view.state.selection;
	if (!empty) return undefined;
	for (const node of [$from.nodeAfter, $from.nodeBefore])
		if (node?.type === schema.nodes.noteReference)
			return {
				kind: node.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
				id: String(node.attrs.id),
			};
	return undefined;
}

/**
 * Word's Convert Notes and "Convert to Endnote/Footnote": moves the selected notes to the other
 * part and retargets their references in one undoable transaction. `reservedIds` are the ids
 * each part holds on disk, so a moved note never takes over a deleted note's element.
 */
export function convertNotesInView(
	view: EditorView,
	model: DocumentModel,
	selection: NoteSelection,
	reservedIds?: Partial<Record<NoteKind, Iterable<string>>>,
): NoteConversionPlan {
	const parts = restoreNoteParts(view.state.doc.attrs.noteParts, model);
	const plan = planNoteConversion(parts, selection, model, reservedIds);
	if (!plan.moved.size) return plan;
	const tr = view.state.tr;
	view.state.doc.descendants((node, pos) => {
		if (node.type !== schema.nodes.noteReference) return true;
		const run = inlineNodeRun(node);
		const next = run && retargetNoteReference(run, plan, model);
		if (next && next !== run) {
			const [replacement] = runToInlineNodes(next, schema);
			if (replacement) tr.setNodeMarkup(pos, undefined, replacement.attrs);
		}
		return false;
	});
	view.dispatch(tr.step(new NotePartsStep(notePartsJson(plan))).scrollIntoView());
	return plan;
}

/** Keeps note reference numbers in document order, as Word renumbers notes after every change. */
export function noteNumberingPlugin(): Plugin {
	return new Plugin({
		appendTransaction(transactions, _oldState, state) {
			if (!transactions.some((transaction) => transaction.docChanged)) return null;
			const counters: Record<string, number> = { footnote: 0, endnote: 0 };
			const formats: Record<NoteKind, string> = {
				footnote: state.doc.attrs.footnoteNumFmt ?? 'decimal',
				endnote: state.doc.attrs.endnoteNumFmt ?? 'lowerRoman',
			};
			const tr = state.tr;
			state.doc.descendants((node, pos) => {
				if (node.type !== schema.nodes.noteReference) return true;
				const kind = node.attrs.kind as NoteKind;
				const number = (counters[kind] ?? 0) + 1;
				counters[kind] = number;
				const label = formatNoteNumber(number, formats[kind]);
				if (node.attrs.number !== number || node.attrs.label !== label)
					tr.setNodeMarkup(pos, undefined, { ...node.attrs, number, label });
				return false;
			});
			return tr.docChanged ? tr : null;
		},
	});
}
