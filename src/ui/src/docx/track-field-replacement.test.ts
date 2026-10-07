// @vitest-environment jsdom
import { afterEach, beforeAll, beforeEach, expect, it, vi } from 'vitest';
import { createDocument, loadDocx, type Paragraph } from 'ooxml-core/docx';
import { collectRevisionRanges } from 'ooxml-core/docx/ui';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory, redo, undo } from 'prosemirror-history';
import { DocxEditorElement, registerDocxEditor } from './index';
import { acceptAllChanges, rejectAllChanges } from './review-commands';

registerDocxEditor();
beforeAll(() => {
	Object.assign(Range.prototype, {
		getClientRects: () => [],
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
beforeEach(() => vi.stubGlobal('ClipboardEvent', Event));
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

for (const input of ['typing', 'paste'] as const)
	for (const mode of ['accept', 'reject'] as const)
		it(`mounted whole-result ${input} records revisions and supports ${mode}, history and export`, async () => {
			const model = createDocument();
			model.trackChanges = true;
			model.blocks = [
				{
					type: 'paragraph',
					id: 'original',
					runs: [
						{ text: 'AB', field: { instr: 'QUOTE AB', simple: true }, fieldInstanceId: 'first' },
						{ text: 'CD', field: { instr: 'QUOTE CD', simple: true }, fieldInstanceId: 'second' },
						{ text: 'R' },
					],
				},
			];
			const element = document.createElement('docx-editor') as DocxEditorElement;
			element.documentModel = model;
			document.body.append(element);
			const view = (element as unknown as { view: EditorView }).view;
			const original = view.state.doc;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 3)));
			if (input === 'paste') expect(view.pasteHTML('<b>X</b>')).toBe(true);
			else {
				let handled = false;
				view.someProp('handleTextInput', (handler) => {
					handled = Boolean(handler(view, 1, 3, 'X', () => view.state.tr.insertText('X', 1, 3)));
					return handled;
				});
				expect(handled).toBe(true);
			}
			const pending = view.state.doc;
			expect(
				collectRevisionRanges(pending)
					.map((range) => range.kind)
					.sort(),
			).toEqual(['delete', 'insert']);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(original)).toBe(true);
			expect(redo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
			view.dispatch(closeHistory(view.state.tr));
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view)).toBe(true);
			expect(view.state.doc.textContent).toBe(mode === 'accept' ? 'XCDR' : 'ABCDR');
			expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			const saved = await loadDocx(await element.saveBytes());
			const paragraph = saved.model.blocks[0] as Paragraph;
			expect(paragraph.runs.map((run) => run.text).join('')).toBe(
				mode === 'accept' ? 'XCDR' : 'ABCDR',
			);
			expect([
				...new Set(paragraph.runs.flatMap((run) => (run.field ? [run.field.instr] : []))),
			]).toEqual(['QUOTE AB', 'QUOTE CD']);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
		});
