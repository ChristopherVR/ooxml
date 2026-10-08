// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { afterEach, expect, it, vi } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { loadDocx, type DocumentModel } from 'ooxml-core/docx';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { wordYjsPluginKey, commentIdsAtSelection, fieldResultRanges } from 'ooxml-core/docx/ui';
import { DocxEditorElement } from './index';
import './index';
import { applyFontFormat } from './font-format';

const editors: DocxEditorElement[] = [];
const sessions: CollabSession[] = [];
const viewOf = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const runsOf = (model: DocumentModel) =>
	model.blocks.flatMap((block) => (block.type === 'paragraph' ? block.runs : []));
afterEach(() => {
	for (const editor of editors.splice(0)) {
		editor.stopCollaboration(true);
		editor.remove();
	}
	for (const session of sessions.splice(0)) session.destroy();
	vi.unstubAllGlobals();
});

// Each case builds two collaborating editors and replays dozens of synced edits, saves
// and reloads in jsdom, about 3 s alone and several times that under full-suite load, so
// the cases carry their own budget instead of the global 20 s timeout.
const CASE_TIMEOUT_MS = 90_000;

for (const name of [
	'picture',
	'note',
	'break',
	'line-break',
	'field',
	'simple-field',
	'adjacent-fields',
])
	it(
		`exports concurrent ${name} comments, deletes independently, and preserves anchors during font edits and detached saves`,
		async () => {
			const bytes = new Uint8Array(
				await readFile(
					name === 'simple-field' || name === 'adjacent-fields'
						? `../core/docx/__fixtures__/field-comments/${name === 'simple-field' ? 'simple' : 'adjacent'}-source.docx`
						: `../core/docx/__fixtures__/review-advanced-object-formatting/${name}-before.docx`,
				),
			);
			let deliver = true;
			const hub = createMemoryHub({ filter: () => deliver });
			for (const author of ['Ada', 'Bob']) {
				sessions.push(
					createCollabSession({
						roomId: 'inline-comments',
						provider: transportProvider({ transport: hub.createTransport('inline-comments') }),
						user: { name: author },
						heartbeatMs: 0,
						teardown: false,
					}),
				);
				const editor = document.createElement('docx-editor') as DocxEditorElement;
				document.body.append(editor);
				editors.push(editor);
				await editor.load(bytes);
			}
			const [a, b] = editors;
			a!.startYjsCollaboration(sessions[0]!, { documentId: 'source', initializeIfEmpty: true });
			b!.startYjsCollaboration(sessions[1]!, { documentId: 'source' });
			const av = viewOf(a!);
			const bv = viewOf(b!);
			const ab = wordYjsPluginKey.getState(av.state)!;
			const bb = wordYjsPluginKey.getState(bv.state)!;
			if (name === 'simple-field' || name === 'adjacent-fields') {
				vi.stubGlobal('ClipboardEvent', Event);
				const field = fieldResultRanges(av.state.doc)[0]!;
				ab.stopCapturing();
				av.dispatch(
					av.state.tr.setSelection(TextSelection.create(av.state.doc, field.from, field.to)),
				);
				const cutData: Record<string, string> = {};
				const cutEvent = new Event('cut', { bubbles: true, cancelable: true });
				Object.defineProperty(cutEvent, 'clipboardData', {
					value: {
						clearData: () => {},
						setData: (kind: string, value: string) => (cutData[kind] = value),
					},
				});
				av.dom.dispatchEvent(cutEvent);
				expect(cutData['text/plain']).toBe('ABCDE');
				expect(cutData['text/html']).not.toContain('data-field');
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(
					runsOf((await loadDocx(await b!.saveBytes())).model).filter(
						(run) => run.fieldChar === 'begin',
					),
				).toHaveLength(1);
				expect(ab.undo()).toBe(true);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(fieldResultRanges(bv.state.doc)[0]!.text).toBe('ABCDE');
				ab.stopCapturing();
				av.dispatch(
					av.state.tr.setSelection(TextSelection.create(av.state.doc, field.from, field.to)),
				);
				av.someProp('handleKeyDown', (handler) =>
					handler(av, new KeyboardEvent('keydown', { key: 'Delete' })),
				);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				const emptySaved = (await loadDocx(await b!.saveBytes())).model;
				expect(runsOf(emptySaved).filter((run) => run.fieldChar === 'begin')).toHaveLength(1);
				expect(
					runsOf(emptySaved)
						.filter((run) => run.field)
						.map((run) => run.text),
				).toEqual(name === 'simple-field' ? [] : ['ABCDE']);
				expect(ab.undo()).toBe(true);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(fieldResultRanges(bv.state.doc)[0]!.text).toBe('ABCDE');
				ab.stopCapturing();
				av.dispatch(
					av.state.tr.setSelection(
						TextSelection.create(av.state.doc, field.from + 1, field.from + 2),
					),
				);
				expect(av.pasteHTML('<b>X</b>')).toBe(true);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(fieldResultRanges(bv.state.doc).map((run) => run.text)).toEqual(
					name === 'simple-field' ? ['AXCDE'] : ['AXCDE', 'ABCDE'],
				);
				const saved = (await loadDocx(await b!.saveBytes())).model;
				expect(
					runsOf(saved)
						.filter((run) => run.field)
						.map((run) => run.fieldInstanceId),
				).toEqual(
					runsOf(a!.documentModel!)
						.filter((run) => run.field)
						.map((run) => run.fieldInstanceId),
				);
				expect(ab.undo()).toBe(true);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(fieldResultRanges(bv.state.doc)[0]!.text).toBe('ABCDE');
				ab.stopCapturing();
				av.dispatch(
					av.state.tr.setSelection(TextSelection.create(av.state.doc, field.from, field.to)),
				);
				const { dom } = av.serializeForClipboard(av.state.selection.content());
				av.dispatch(av.state.tr.setSelection(TextSelection.create(av.state.doc, field.to)));
				expect(av.pasteHTML(dom.innerHTML)).toBe(true);
				expect(av.state.doc.eq(bv.state.doc)).toBe(true);
				expect(fieldResultRanges(bv.state.doc).map((run) => run.text)).toEqual(
					name === 'simple-field' ? ['ABCDE'] : ['ABCDE', 'ABCDE'],
				);
			}
			let pos = -1;
			av.state.doc.descendants((node, position) => {
				if (
					pos < 0 &&
					(name === 'field' || name === 'simple-field' || name === 'adjacent-fields'
						? node.isText && node.marks.some((mark) => mark.type.name === 'field')
						: node.isInline && !node.isText)
				)
					pos = position;
			});
			for (const view of [av, bv])
				view.dispatch(
					view.state.tr.setSelection(TextSelection.create(view.state.doc, pos, pos + 1)),
				);
			deliver = false;
			expect(ab.comments.add(av, 'Ada', 'A comment', () => 'a')).not.toBeNull();
			expect(bb.comments.add(bv, 'Bob', 'B comment', () => 'b')).not.toBeNull();
			deliver = true;
			sessions[0]!.resync();
			expect(av.state.doc.eq(bv.state.doc)).toBe(true);
			for (const editor of editors) {
				expect(commentIdsAtSelection(viewOf(editor))).toEqual(['a', 'b']);
				expect(runsOf(editor.documentModel!).flatMap((run) => run.commentIds ?? [])).toHaveLength(
					name === 'field' ? 10 : 2,
				);
				const model = (await loadDocx(await editor.saveBytes())).model;
				expect(model.comments?.map((comment) => comment.text).sort()).toEqual([
					'A comment',
					'B comment',
				]);
				const runs = runsOf(model);
				if (name === 'adjacent-fields') {
					const fields = runs.filter((run) => run.field);
					expect(fields).toHaveLength(2);
					expect(fields[0]!.commentIds).toHaveLength(2);
					expect(fields[1]!.commentIds).toBeUndefined();
				}
				expect(runs.flatMap((run) => run.commentIds ?? [])).toHaveLength(name === 'field' ? 10 : 2);
			}
			bb.stopCapturing();
			applyFontFormat(bv, { size: 18, smallCaps: true });
			expect(commentIdsAtSelection(av)).toEqual(['a', 'b']);
			expect(bb.undo()).toBe(true);
			expect(commentIdsAtSelection(av)).toEqual(['a', 'b']);
			deliver = false;
			expect(ab.comments.delete(av, 'a')).toBe(true);
			expect(bb.comments.delete(bv, 'b')).toBe(true);
			deliver = true;
			sessions[0]!.resync();
			expect(commentIdsAtSelection(av)).toEqual([]);
			expect(runsOf(a!.documentModel!).every((run) => !run.commentIds?.length)).toBe(true);
			const deleted = (await loadDocx(await a!.saveBytes())).model;
			expect(deleted.comments ?? []).toEqual([]);
			expect(runsOf(deleted).every((run) => !run.commentIds?.length)).toBe(true);
			expect(ab.undo()).toBe(true);
			expect(commentIdsAtSelection(bv)).toEqual(['a']);
			a!.remove();
			expect(bb.comments.reply('a', 'Bob', 'Remote reply', () => 'reply')).toBe(true);
			const detached = (await loadDocx(await a!.saveBytes())).model;
			expect(detached.comments).toHaveLength(2);
			expect(runsOf(detached).flatMap((run) => run.commentIds ?? [])).toHaveLength(
				name === 'field' ? 5 : 1,
			);
		},
		CASE_TIMEOUT_MS,
	);
