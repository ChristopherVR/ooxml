import {
	glossaryEntryKey,
	documentPropertyChoice,
	documentPropertyControl,
	documentPropertyValue,
	glossaryBlocksCopy,
	type Block,
	type DocumentModel,
	type DocumentPropertyId,
	type GlossaryEntry,
} from 'ooxml-core/docx';
import { everyControlSpan, insertBoundControl } from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import { Fragment, Slice } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import type { EditorLocale } from './localization';
import './locales/insert-parts-translations';
import { translate } from './localization';
import type { QuickPartCommand } from './ribbon-action';

let counter = 0;

/**
 * Quick Parts > Document Property: a control bound to the property, showing its value (or, while
 * the property is empty, its name in brackets) and saved back to the document's property part.
 */
export function insertDocumentProperty(
	view: EditorView,
	model: DocumentModel,
	id: DocumentPropertyId,
	locale: EditorLocale,
): boolean {
	if (!view.editable) return false;
	const used = new Set(everyControlSpan(view.state.doc).map((span) => span.control.id));
	const label = translate(
		locale,
		(documentPropertyChoice(id)?.label ?? id) as Parameters<typeof translate>[1],
	);
	const made = documentPropertyControl(id, used, label);
	const choice = documentPropertyChoice(id);
	if (!made || !choice) return false;
	const tr = insertBoundControl(
		view.state,
		{ ...made.control, alias: label, tag: choice.label },
		documentPropertyValue(model, id),
		`[${label}]`,
	);
	if (!tr) return false;
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}

/**
 * Inserts a glossary entry at the selection. A one-paragraph entry goes in as text at the cursor,
 * as Word inserts such a building block; a longer one goes in as paragraphs.
 */
export function insertGlossaryEntry(
	view: EditorView,
	model: DocumentModel,
	entry: GlossaryEntry,
): boolean {
	if (!view.editable || !entry.blocks.length) return false;
	const blocks = glossaryBlocksCopy(entry.blocks, `qp${Date.now().toString(36)}-${counter++}-`);
	const content = modelToDoc({ ...model, blocks }).content;
	const single = content.childCount === 1 && content.firstChild?.type.name === 'paragraph';
	const slice = single ? new Slice(content, 1, 1) : new Slice(content, 0, 0);
	view.dispatch(closeHistory(view.state.tr.replaceSelection(slice)).scrollIntoView());
	return true;
}

/** The selected content as model blocks (for saving it to a gallery), or undefined when it is empty. */
export function selectionBlocks(view: EditorView, model: DocumentModel): Block[] | undefined {
	const { selection } = view.state;
	if (selection.empty) return undefined;
	try {
		const slice = selection.content();
		const doc = view.state.doc.type.create(view.state.doc.attrs, Fragment.from(slice.content));
		const blocks = docToModel(doc, model).blocks;
		return blocks.length ? blocks : undefined;
	} catch {
		return undefined;
	}
}

/** Insert > Quick Parts: a document property or a gallery entry, at the selection. */
export async function runQuickPart(
	view: EditorView,
	model: DocumentModel,
	command: QuickPartCommand,
	locale: EditorLocale,
): Promise<void> {
	if (command.kind === 'property') insertDocumentProperty(view, model, command.id, locale);
	else {
		const entry = model.glossary?.find((item) => glossaryEntryKey(item) === command.key);
		if (entry) insertGlossaryEntry(view, model, entry);
	}
	view.focus();
}
