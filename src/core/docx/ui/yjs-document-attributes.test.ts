// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import * as Y from 'yjs';
import { ySyncPluginKey } from 'y-prosemirror';
import { createCollabSession, createMemoryHub, transportProvider } from '../../collab/index';
import { WordYjsCollaboration } from './yjs-collaboration';
import { markSpecs } from './schema-marks';
import { hardBreakNodeSpec } from './break-note-schema';
import { commentHistoryAttributes } from './comment-history';
const schema = new Schema({
	nodes: {
		doc: {
			content: 'paragraph+',
			attrs: { pageWidth: { default: 816 }, ...commentHistoryAttributes },
		},
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
		hardBreak: hardBreakNodeSpec,
	},
	marks: markSpecs,
});
for (const order of ['settings-first', 'body-first'])
	for (const viewer of [false, true])
		it(`retains ${viewer ? 'read-only' : 'editing'} peer body edits when updates arrive ${order}`, () => {
			const hub = createMemoryHub();
			const sessions = [0, 1].map((index) =>
				createCollabSession({
					roomId: 'attributes',
					provider: transportProvider({ transport: hub.createTransport('attributes') }),
					user: { name: String(index), ...(viewer && index ? { role: 'viewer' as const } : {}) },
					heartbeatMs: 0,
					teardown: false,
				}),
			);
			const initial = schema.node(
				'doc',
				{ commentThreads: '[]' },
				schema.node('paragraph', null, schema.text('Alpha')),
			);
			const bindings = sessions.map(
				(session, index) =>
					new WordYjsCollaboration(session, initial, {
						documentId: 'source',
						initializeIfEmpty: index === 0,
					}),
			);
			const views = bindings.map(
				(binding) =>
					new EditorView(document.createElement('div'), {
						state: EditorState.create(binding.state(schema)),
					}),
			);
			try {
				const doc = sessions[0]!.doc;
				const attrs = doc.getMap('docx:attributes');
				expect(attrs.has('commentThreads')).toBe(false);
				const text = (bindings[0]!.fragment.get(0) as Y.XmlElement).get(0) as Y.XmlText;
				doc.transact(() => {
					if (order === 'settings-first') attrs.set('pageWidth', 900);
					text.insert(0, 'New ');
					if (order === 'body-first') attrs.set('pageWidth', 900);
				}, ySyncPluginKey);
				for (const view of views) {
					expect(view.state.doc.textContent).toBe('New Alpha');
					expect(view.state.doc.attrs.pageWidth).toBe(900);
					expect(view.state.doc.attrs.commentThreads).toBeNull();
				}
				expect(views[0]!.state.doc.eq(views[1]!.state.doc)).toBe(true);
				expect(bindings[0]!.undo()).toBe(true);
				for (const view of views) {
					expect(view.state.doc.textContent).toBe('Alpha');
					expect(view.state.doc.attrs.pageWidth).toBe(816);
				}
			} finally {
				for (const view of views) view.destroy();
				for (const binding of bindings) binding.destroy();
				for (const session of sessions) session.destroy();
			}
		});
