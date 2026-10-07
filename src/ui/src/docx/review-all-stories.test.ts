// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { loadDocx, listRevisions, type DocumentModel, type Block } from 'ooxml-core/docx';
import { undo, redo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement } from './index';
import './index';

function content(model: DocumentModel) {
	const blocks = (items: Block[]): unknown[] =>
		items.map((block) =>
			block.type === 'paragraph'
				? block.runs.map((run) => ({
						text: run.text,
						bold: run.bold ?? false,
						italic: run.italic ?? false,
						noteReference: run.noteReference ?? null,
						noteMark: run.noteMark ?? null,
					}))
				: block.rows.map((row) => row.map((cell) => blocks(cell.paragraphs))),
		);
	return {
		body: blocks(model.blocks),
		headers: model.sections?.map((section) => blocks(section.headers?.default?.blocks ?? [])),
		footers: model.sections?.map((section) => blocks(section.footers?.default?.blocks ?? [])),
		footnotes: model.footnotes?.map((note) => blocks(note.blocks)),
		endnotes: model.endnotes?.map((note) => blocks(note.blocks)),
	};
}
for (const mode of ['accept', 'reject'] as const)
	it(`${mode}s native formatting in all five stories and restores them with one undo`, async () => {
		const directory = resolve('../core/docx/__fixtures__/review-stories');
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		try {
			await editor.load(
				new Uint8Array(await readFile(resolve(directory, 'all-stories-tracked.docx'))),
			);
			const source = structuredClone(editor.documentModel!);
			expect(listRevisions(source)).toHaveLength(5);
			editor
				.shadowRoot!.querySelector<HTMLButtonElement>(
					`[aria-label="${mode === 'accept' ? 'Accept' : 'Reject'} all"]`,
				)!
				.click();
			expect(listRevisions(editor.documentModel!)).toHaveLength(0);
			const actual = (await loadDocx(await editor.saveBytes())).model;
			const native = (
				await loadDocx(
					await readFile(
						resolve(directory, `all-stories-${mode === 'accept' ? 'accepted' : 'rejected'}.docx`),
					),
				)
			).model;
			expect(listRevisions(actual)).toHaveLength(0);
			expect(content(actual)).toEqual(content(native));
			const view = (editor as unknown as { view: EditorView }).view;
			const resolved = view.state.doc;
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(editor.documentModel).toEqual(source);
			expect(redo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(resolved)).toBe(true);
		} finally {
			editor.remove();
		}
	});
