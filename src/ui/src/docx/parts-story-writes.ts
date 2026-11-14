// The body transactions that store a header, footer or note's new blocks. Story text is kept in
// the body document's attributes, so these edits share the body's undo history.
import type { Block, DocumentModel, HeaderFooterContent, Note } from 'ooxml-core/docx';
import { sectionLayoutJson } from 'ooxml-core/docx';
import {
	notePartsJson,
	NotePartsStep,
	sectionPartsJson,
	SectionPartsStep,
} from 'ooxml-core/docx/ui';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { HeaderFooterSlotName } from './header-footer-editor';
import { effectiveHeaderFooter } from './header-footer-link';
import { newHeaderFooterId, withBlankHeaderFooter } from './header-footer-commands';
import type { NoteKind } from './note-commands';

/** Replaces footnotes' or endnotes' blocks, by note id. */
export function notesTransaction(
	state: EditorState,
	model: DocumentModel,
	kind: NoteKind,
	blocks: ReadonlyMap<string, Block[]>,
): Transaction {
	const replace = (notes: Note[]) =>
		notes.map((note) => {
			const next = blocks.get(note.id);
			return next ? { ...note, blocks: structuredClone(next) } : note;
		});
	const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
	const next = { ...model, ...(model[key] && { [key]: replace(model[key]) }) };
	return state.tr.step(new NotePartsStep(notePartsJson(next)));
}

/** Applies header/footer blocks to every section slot that shares the edited part. */
export function headerFooterTransaction(
	state: EditorState,
	source: DocumentModel,
	kind: 'headers' | 'footers',
	slot: HeaderFooterSlotName,
	sectionIndex: number,
	blocks: Block[],
): Transaction {
	let model = source;
	if (!effectiveHeaderFooter(model, sectionIndex, kind, slot))
		model = withBlankHeaderFooter(model, kind, newHeaderFooterId, sectionIndex, slot);
	const edited = effectiveHeaderFooter(model, sectionIndex, kind, slot)!;
	const sections = (model.sections ?? []).map((section) => {
		const slots = { ...section[kind] };
		for (const [name, content] of Object.entries(slots) as [
			HeaderFooterSlotName,
			HeaderFooterContent,
		][])
			if (content === edited || (edited.partName && content?.partName === edited.partName))
				slots[name] = { ...content, blocks: structuredClone(blocks) };
		return { ...section, [kind]: slots };
	});
	const tr = state.tr;
	const layout = sectionLayoutJson(sections);
	if (layout !== state.doc.attrs.sections) tr.setDocAttribute('sections', layout);
	if (Boolean(model.evenAndOddHeaders) !== state.doc.attrs.evenAndOddHeaders)
		tr.setDocAttribute('evenAndOddHeaders', Boolean(model.evenAndOddHeaders));
	return tr.step(new SectionPartsStep(sectionPartsJson(sections)));
}

/** The story key (`findStoryKey`) of the header or footer a section shows in `slot`. */
export function headerFooterStoryKey(
	model: DocumentModel,
	kind: 'headers' | 'footers',
	slot: HeaderFooterSlotName,
	sectionIndex: number,
): string {
	const sections = model.sections ?? [];
	let owner = sectionIndex;
	while (owner > 0 && !sections[owner]?.[kind]?.[slot]) owner--;
	const partName = sections[owner]?.[kind]?.[slot]?.partName;
	const first = partName
		? sections.findIndex((section) => section[kind]?.[slot]?.partName === partName)
		: -1;
	if (first >= 0 && first < owner) owner = first;
	return `${kind === 'headers' ? 'header' : 'footer'}:${owner}:${slot}`;
}
