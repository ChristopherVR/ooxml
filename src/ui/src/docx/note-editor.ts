import type { Block, Note, Paragraph, TextRun } from 'ooxml-core/docx';
import { translateUiText } from './localization';
import { openBlocksEditor, type InlineEditorOptions } from './header-footer-editor';
import type { NoteKind } from './note-commands';

export interface NoteEditingOptions {
	editor?: InlineEditorOptions;
	note(kind: NoteKind, id: string): Note | undefined;
	change(kind: NoteKind, id: string, blocks: Block[]): void;
	editable(): boolean;
}

/** Splits off the note's automatic number mark so it can't be edited away. */
function withoutMark(blocks: Block[]): { editable: Block[]; mark?: TextRun } {
	const [first, ...rest] = blocks;
	if (first?.type !== 'paragraph' || !first.runs[0]?.noteMark) return { editable: blocks };
	const [mark, ...runs] = first.runs;
	return {
		editable: [{ ...first, runs: runs.length ? runs : [{ text: '' }] }, ...rest],
		mark,
	};
}

function withMark(blocks: Block[], mark: TextRun | undefined): Block[] {
	const [first, ...rest] = blocks;
	if (!mark || first?.type !== 'paragraph') return blocks;
	const paragraph: Paragraph = { ...first, runs: [mark, ...first.runs] };
	return [paragraph, ...rest];
}

/** Double-click a footnote or endnote to edit its text in place. */
export function attachNoteEditing(root: HTMLElement, options: NoteEditingOptions): void {
	for (const item of root.querySelectorAll<HTMLElement>('li[data-docx-note-id]')) {
		const id = item.dataset.docxNoteId!;
		const kind = item.dataset.docxNoteKind;
		if (kind !== 'footnote' && kind !== 'endnote') continue;
		const body = item.querySelector<HTMLElement>('.dve-note-body');
		if (!body) continue;
		item.title = translateUiText(root, 'Double-click to edit');
		item.addEventListener('dblclick', () => {
			if (!options.editable() || item.classList.contains('dve-header-footer-editing')) return;
			const note = options.note(kind, id);
			if (!note) return;
			const { editable, mark } = withoutMark(note.blocks);
			openBlocksEditor(
				item,
				body,
				editable,
				(blocks) => options.change(kind, id, withMark(blocks, mark)),
				options.editor,
			);
		});
	}
}
