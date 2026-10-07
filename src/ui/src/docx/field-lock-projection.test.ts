// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it } from 'vitest';
import { createDocument, loadDocx, type Paragraph } from 'ooxml-core/docx';
import { collectRevisionRanges, inlineNodeRun } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { closeHistory, undo, redo } from 'prosemirror-history';
import { DocxEditorElement, registerDocxEditor } from './index';
import { acceptAllChanges, rejectAllChanges } from './review-commands';
import { updateFields } from './field-update';

registerDocxEditor();
beforeAll(() => {
	Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
	Object.assign(Range.prototype, {
		getClientRects: () => [],
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => document.body.replaceChildren());

function mount(tracked: boolean, formatting: boolean, text = '') {
	const model = createDocument();
	model.trackChanges = tracked;
	model.trackFormatting = formatting;
	model.blocks = [
		{
			type: 'paragraph',
			id: 'field',
			runs: [
				{ text: '', fieldChar: 'begin', fieldFlags: { locked: true, dirty: false } },
				{ text: '', fieldCode: 'REF Target' },
				{ text: '', fieldChar: 'separate' },
				...(text ? [{ text, field: { instr: 'REF Target' }, fieldFlags: { locked: true } }] : []),
				{ text: '', fieldChar: 'end' },
			],
		},
		{ type: 'paragraph', id: 'target', bookmarks: ['Target'], runs: [{ text: 'Replacement' }] },
	];
	const element = document.createElement('docx-editor') as DocxEditorElement;
	element.documentModel = model;
	document.body.append(element);
	return { element, view: (element as unknown as { view: EditorView }).view };
}

for (const formatting of [true, false])
	for (const action of ['accept', 'reject'] as const)
		it(`mounted empty locked result typing records once, supports ${action}, undo and save (format tracking ${formatting})`, async () => {
			const { element, view } = mount(true, formatting);
			const original = view.state.doc;
			view.dispatch(view.state.tr.insertText('New', 4));
			const pending = view.state.doc;
			expect(collectRevisionRanges(pending).map((range) => range.kind)).toEqual(['insert']);
			expect(inlineNodeRun(pending.nodeAt(4)!)).toMatchObject({
				text: 'New',
				fieldFlags: { locked: true },
				revision: { kind: 'insert' },
			});
			expect(updateFields(view, () => undefined)).toBe(false);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(original)).toBe(true);
			expect(redo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
			view.dispatch(closeHistory(view.state.tr));
			expect((action === 'accept' ? acceptAllChanges : rejectAllChanges)(view)).toBe(true);
			expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			const saved = await loadDocx(await element.saveBytes());
			const runs = (saved.model.blocks[0] as Paragraph).runs;
			expect(runs.find((run) => run.fieldChar === 'begin')?.fieldFlags).toEqual({
				locked: true,
				dirty: false,
			});
			if (action === 'accept')
				expect(runs.find((run) => run.text === 'New')?.fieldFlags?.locked).toBe(true);
			else expect(runs.some((run) => run.text === 'New')).toBe(false);
		});

it('mounted delete and retype keeps a locked REF cached through update and export', async () => {
	const { element, view } = mount(false, true, 'Old');
	view.dispatch(view.state.tr.delete(4, 7));
	view.dispatch(view.state.tr.insertText('New', 4));
	expect(inlineNodeRun(view.state.doc.nodeAt(4)!)).toMatchObject({
		text: 'New',
		fieldFlags: { locked: true },
	});
	expect(updateFields(view, () => undefined)).toBe(false);
	const saved = await loadDocx(await element.saveBytes());
	expect(
		(saved.model.blocks[0] as Paragraph).runs.find((run) => run.text === 'New')?.fieldFlags?.locked,
	).toBe(true);
});
