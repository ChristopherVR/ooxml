// @vitest-environment jsdom
import { afterEach, beforeAll, expect, it } from 'vitest';
import { createDocument, loadDocx, sectionsOf } from 'ooxml-core/docx';
import { collectRevisionRanges, wordYjsPluginKey } from 'ooxml-core/docx/ui';
import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory, undo, redo } from 'prosemirror-history';
import { DocxEditorElement, registerDocxEditor } from './index';
import { acceptAllChanges, rejectAllChanges } from './review-commands';

registerDocxEditor();
beforeAll(() => {
	Object.defineProperty(document, 'execCommand', { configurable: true, value: () => false });
	Object.assign(Range.prototype, {
		getClientRects: () => [],
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => document.body.replaceChildren());

function mount(action: 'split' | 'join') {
	const model = createDocument();
	model.trackChanges = true;
	model.blocks =
		action === 'split'
			? [
					{
						type: 'paragraph',
						id: 'original',
						align: 'center',
						runs: [{ text: 'Hello', bold: true }],
					},
				]
			: [
					{ type: 'paragraph', id: 'first', align: 'center', runs: [{ text: 'He', bold: true }] },
					{ type: 'paragraph', id: 'last', align: 'center', runs: [{ text: 'llo', bold: true }] },
				];
	const element = document.createElement('docx-editor') as DocxEditorElement;
	model.sections = sectionsOf(model);
	element.documentModel = model;
	document.body.append(element);
	return { element, view: (element as unknown as { view: EditorView }).view };
}

for (const action of ['split', 'join'] as const)
	for (const mode of ['accept', 'reject'] as const)
		it(`mounted ${action === 'split' ? 'Enter' : 'Backspace'} records a paragraph mark, supports ${mode}, undo and model projection`, () => {
			const { element, view } = mount(action);
			const original = view.state.doc;
			const pos = action === 'split' ? 3 : view.state.doc.firstChild!.nodeSize + 1;
			view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
			view.focus();
			const key = new KeyboardEvent('keydown', {
				key: action === 'split' ? 'Enter' : 'Backspace',
				code: action === 'split' ? 'Enter' : 'Backspace',
				keyCode: action === 'split' ? 13 : 8,
				bubbles: true,
				cancelable: true,
			});
			view.dom.dispatchEvent(key);
			expect(key.defaultPrevented).toBe(true);
			const pending = view.state.doc;
			expect(pending.childCount).toBe(2);
			expect(collectRevisionRanges(pending)).toHaveLength(1);
			expect(pending.firstChild!.attrs.markRevision?.kind).toBe(
				action === 'split' ? 'insert' : 'delete',
			);
			expect(
				new Set(
					Array.from({ length: pending.childCount }, (_, index) => pending.child(index).attrs.id),
				).size,
			).toBe(2);
			expect(element.documentModel!.blocks[0]).toMatchObject({
				markRevision: { kind: action === 'split' ? 'insert' : 'delete' },
			});
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(original)).toBe(true);
			expect(redo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
			view.dispatch(closeHistory(view.state.tr));
			expect((mode === 'accept' ? acceptAllChanges : rejectAllChanges)(view)).toBe(true);
			const keep = (mode === 'accept') === (action === 'split');
			expect(view.state.doc.childCount).toBe(keep ? 2 : 1);
			expect(collectRevisionRanges(view.state.doc)).toEqual([]);
			expect(
				element.documentModel!.blocks.every(
					(block) => block.type === 'paragraph' && block.align === 'center' && !block.markRevision,
				),
			).toBe(true);
			expect(undo(view.state, view.dispatch)).toBe(true);
			expect(view.state.doc.eq(pending)).toBe(true);
		});

it('mounted Backspace removes an own pending Enter boundary with original identity', () => {
	const { view } = mount('split');
	const original = view.state.doc;
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
	view.dom.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true, cancelable: true }),
	);
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, view.state.doc.firstChild!.nodeSize + 1),
		),
	);
	view.dom.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Backspace', keyCode: 8, bubbles: true, cancelable: true }),
	);
	expect(view.state.doc.eq(original)).toBe(true);
	expect(collectRevisionRanges(view.state.doc)).toEqual([]);
});

it.each(['split', 'join'] as const)(
	'converges a tracked plain %s and local undo through two Yjs peers',
	async (action) => {
		const hub = createMemoryHub();
		const sessions = ['Ada', 'Bob'].map((name) =>
			createCollabSession({
				roomId: 'paragraph-boundaries',
				provider: transportProvider({ transport: hub.createTransport('paragraph-boundaries') }),
				user: { name },
				heartbeatMs: 0,
				teardown: false,
			}),
		);
		const a = mount(action);
		const b = mount(action);
		try {
			a.element.startYjsCollaboration(sessions[0]!, {
				documentId: 'source',
				initializeIfEmpty: true,
			});
			b.element.startYjsCollaboration(sessions[1]!, { documentId: 'source' });
			const av = (a.element as unknown as { view: EditorView }).view;
			const bv = (b.element as unknown as { view: EditorView }).view;
			const binding = wordYjsPluginKey.getState(av.state)!;
			const original = av.state.doc;
			binding.stopCapturing();
			const pos = action === 'split' ? 3 : av.state.doc.firstChild!.nodeSize + 1;
			av.dispatch(av.state.tr.setSelection(TextSelection.create(av.state.doc, pos)));
			av.dom.dispatchEvent(
				new KeyboardEvent('keydown', {
					key: action === 'split' ? 'Enter' : 'Backspace',
					keyCode: action === 'split' ? 13 : 8,
					bubbles: true,
					cancelable: true,
				}),
			);
			expect(av.state.doc.eq(bv.state.doc)).toBe(true);
			expect(collectRevisionRanges(av.state.doc)).toHaveLength(1);
			expect(collectRevisionRanges(bv.state.doc)).toEqual(collectRevisionRanges(av.state.doc));
			const saved = await loadDocx(await b.element.saveBytes());
			expect(saved.model.blocks[0]).toMatchObject({
				markRevision: { kind: action === 'split' ? 'insert' : 'delete' },
			});
			expect(binding.undo()).toBe(true);
			expect(av.state.doc.eq(original)).toBe(true);
			expect(av.state.doc.eq(bv.state.doc)).toBe(true);
			expect(collectRevisionRanges(bv.state.doc)).toEqual([]);
		} finally {
			a.element.stopCollaboration(true);
			b.element.stopCollaboration(true);
			for (const session of sessions) session.destroy();
		}
	},
);
