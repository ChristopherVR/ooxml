/**
 * Pasting from Word: its clipboard HTML is normalised by the core reader before ProseMirror
 * parses it, so `mso-list` paragraphs paste as real lists (each pasted list gets a new
 * definition in the document's numbering), classes and headings as the document's own
 * paragraph styles, and Word-only run formatting as run properties. Other HTML passes through.
 */
import {
	createListDefinition,
	normalizeWordClipboardHtml,
	styleIdForName,
	type DocumentModel,
	type WordPasteMode,
} from 'ooxml-core/docx';
import { Slice } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';

/** Word's HTML in the editor's terms, registering its lists in `model`; `null` when not Word's. */
export function wordHtmlForPaste(
	html: string,
	model: DocumentModel,
	mode: WordPasteMode = 'source',
): string | null {
	return normalizeWordClipboardHtml(html, {
		mode,
		numId: (list) => {
			try {
				const created = createListDefinition(model.numberingCatalog, list.levels);
				model.numberingCatalog = created.catalog;
				return created.numId;
			} catch {
				// A definition the model would reject pastes as plain paragraphs, never a broken list.
				return undefined;
			}
		},
		paragraphStyleId: (name) => styleIdForName(model.paragraphStyles?.styles, name),
		characterStyleId: (name) => styleIdForName(model.characterStyles?.styles, name),
	});
}

export const wordPasteKey = new PluginKey('dveWordPaste');

/** Normalises pasted Word HTML for the body editor; `mode` is read at each paste. */
export function wordPastePlugin(
	model: () => DocumentModel,
	mode: () => WordPasteMode = () => 'source',
): Plugin {
	let fromWord = false;
	return new Plugin({
		key: wordPasteKey,
		props: {
			transformPastedHTML: (html) => {
				const normalized = wordHtmlForPaste(html, model(), mode());
				fromWord = normalized !== null;
				return normalized ?? html;
			},
			transformPasted: (slice, view) => {
				const word = fromWord;
				fromWord = false;
				// Into an empty paragraph, Word's first pasted paragraph keeps its own list and
				// style: the slice starts with a whole paragraph instead of merging into it.
				const { $from, empty } = view.state.selection;
				const first = slice.content.firstChild;
				if (
					!word ||
					!empty ||
					slice.openStart === 0 ||
					$from.parent.type.name !== 'paragraph' ||
					$from.parent.content.size !== 0 ||
					first?.type.name !== 'paragraph'
				)
					return slice;
				return new Slice(slice.content, 0, slice.openEnd);
			},
		},
	});
}
