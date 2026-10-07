// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { afterEach, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { loadDocx, listRevisions, type DocumentModel } from 'ooxml-core/docx';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import {
	toggleTrackChanges,
	wordYjsPluginKey,
	collectRevisionRanges,
	rejectAllChanges,
	acceptAllChanges,
} from 'ooxml-core/docx/ui';
import { DocxEditorElement } from './index';
import './index';
import { applyFontFormat } from './font-format';

const editors: DocxEditorElement[] = [];
const sessions: CollabSession[] = [];
afterEach(() => {
	for (const editor of editors.splice(0)) {
		editor.stopCollaboration(true);
		editor.remove();
	}
	for (const session of sessions.splice(0)) session.destroy();
});
const viewOf = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const content = (model: DocumentModel) =>
	model.blocks.map((p) =>
		p.type === 'paragraph'
			? p.runs.map(
					({ sourceRunPropertiesXml: _source, restoredRunPropertiesXml: _restored, ...run }) => run,
				)
			: '[table]',
	);

for (const name of ['picture', 'note', 'break', 'line-break'])
	it(`shares one advanced ${name} Font-dialog edit, export, rejection and local undo`, async () => {
		const fixture = (mode: string) =>
			readFile(
				resolve(`../core/docx/__fixtures__/review-advanced-object-formatting/${name}-${mode}.docx`),
			);
		const bytes = new Uint8Array(await fixture('before'));
		const hub = createMemoryHub();
		for (const author of ['Ada', 'Bob']) {
			const session = createCollabSession({
				roomId: 'inline-font',
				provider: transportProvider({ transport: hub.createTransport('inline-font') }),
				user: { name: author },
				heartbeatMs: 0,
				teardown: false,
			});
			sessions.push(session);
			const editor = document.createElement('docx-editor') as DocxEditorElement;
			document.body.append(editor);
			editors.push(editor);
			await editor.load(bytes);
			editor.reviewAuthor = author;
		}
		const [a, b] = editors;
		a!.startYjsCollaboration(sessions[0]!, { documentId: 'source', initializeIfEmpty: true });
		b!.startYjsCollaboration(sessions[1]!, { documentId: 'source' });
		const av = viewOf(a!);
		const bv = viewOf(b!);
		toggleTrackChanges(av.state, av.dispatch, av);
		const collab = wordYjsPluginKey.getState(bv.state)!;
		collab.stopCapturing();
		const initial = bv.state.doc;
		let pos = -1;
		initial.descendants((node, position) => {
			if (pos < 0 && node.isInline && !node.isText) pos = position;
		});
		bv.dispatch(bv.state.tr.setSelection(TextSelection.create(initial, pos, pos + 1)));
		applyFontFormat(bv, {
			size: 18,
			color: '#C00000',
			smallCaps: true,
			spacing: 1.5,
			scale: 150,
			position: 2,
			kerning: 12,
		});
		const tracked = bv.state.doc;
		expect(av.state.doc.eq(tracked)).toBe(true);
		expect(collectRevisionRanges(tracked)).toMatchObject([{ kind: 'formatChange', author: 'Bob' }]);
		for (const editor of editors)
			expect(listRevisions((await loadDocx(await editor.saveBytes())).model)).toHaveLength(1);
		expect(collab.undo()).toBe(true);
		expect(av.state.doc.eq(initial)).toBe(true);
		expect(collab.redo()).toBe(true);
		expect(av.state.doc.eq(tracked)).toBe(true);
		wordYjsPluginKey.getState(av.state)!.stopCapturing();
		expect(acceptAllChanges(av)).toBe(true);
		const accepted = (await loadDocx(await fixture('accepted'))).model;
		for (const editor of editors)
			expect(content((await loadDocx(await editor.saveBytes())).model)).toEqual(content(accepted));
		expect(wordYjsPluginKey.getState(av.state)!.undo()).toBe(true);
		expect(bv.state.doc.eq(tracked)).toBe(true);
		wordYjsPluginKey.getState(av.state)!.stopCapturing();
		expect(rejectAllChanges(av)).toBe(true);
		const native = (await loadDocx(await fixture('rejected'))).model;
		for (const editor of editors)
			expect(content((await loadDocx(await editor.saveBytes())).model)).toEqual(content(native));
		expect(wordYjsPluginKey.getState(av.state)!.undo()).toBe(true);
		expect(bv.state.doc.eq(tracked)).toBe(true);
	});
