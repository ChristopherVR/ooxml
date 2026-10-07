import type { DocumentModel } from '../model';
import { createDocument } from '../model';
import { acceptAllRevisions, listRevisions, rejectAllRevisions } from '../revision-commands';
import type { Node } from 'prosemirror-model';
import { DocAttrStep, StepMap } from 'prosemirror-transform';
import { inlineNodeRun } from './run-adapter';

type NoteParts = Pick<DocumentModel, 'footnotes' | 'endnotes'>;

/** Note content is document state, separate from the body and its reference nodes. */
export function notePartsJson(model: NoteParts): string {
	return JSON.stringify({
		...(model.footnotes ? { footnotes: model.footnotes } : {}),
		...(model.endnotes ? { endnotes: model.endnotes } : {}),
	});
}

export function restoreNoteParts(value: unknown, prior: NoteParts): NoteParts {
	return typeof value === 'string' ? (JSON.parse(value) as NoteParts) : prior;
}

/** An identity boundary allows consecutive note edits to share one body-history event. */
export class NotePartsStep extends DocAttrStep {
	constructor(value: string) {
		super('noteParts', value);
	}
	override getMap(): StepMap {
		return new StepMap([0, 0, 0]);
	}
	override invert(doc: Node): NotePartsStep {
		return new NotePartsStep(doc.attrs.noteParts);
	}
}

function notesModel(doc: Node): DocumentModel {
	return { ...createDocument(), blocks: [], ...restoreNoteParts(doc.attrs.noteParts, {}) };
}

export function hasNoteRevisions(doc: Node): boolean {
	return listRevisions(notesModel(doc)).length > 0;
}

function references(doc: Node): Set<string> {
	const ids = new Set<string>();
	doc.descendants((node) => {
		if (!node.isInline) return;
		const reference = inlineNodeRun(node)?.noteReference;
		if (reference) ids.add(`${reference.kind}:${reference.id}`);
	});
	return ids;
}

/** Resolve before dispatch so a failed note restoration cannot partially resolve the body. */
export function resolveNoteRevisions(
	doc: Node,
	resolvedBody: Node,
	mode: 'accept' | 'reject',
): NotePartsStep | undefined {
	const model = notesModel(doc);
	const before = references(doc);
	const after = references(resolvedBody);
	for (const [key, kind] of [
		['footnotes', 'footnote'],
		['endnotes', 'endnote'],
	] as const) {
		if (model[key])
			model[key] = model[key].filter(
				(note) => !before.has(`${kind}:${note.id}`) || after.has(`${kind}:${note.id}`),
			);
	}
	const next = (mode === 'accept' ? acceptAllRevisions : rejectAllRevisions)(model);
	const value = notePartsJson(next);
	return value !== notePartsJson(notesModel(doc)) ? new NotePartsStep(value) : undefined;
}
