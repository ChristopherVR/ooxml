// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from '../../collab/index';
import type { TextRun } from '../model';
import { WordYjsCollaboration } from './yjs-collaboration';
import { imageNodeSpec } from './inline-content-schema';
import { hardBreakNodeSpec, pageBreakNodeSpec, noteReferenceNodeSpec } from './break-note-schema';
import { markSpecs } from './schema-marks';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
import { applyRunFormattingPatch } from './run-format-command';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
		image: imageNodeSpec,
		hardBreak: hardBreakNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
	},
	marks: markSpecs,
});
const sessions: CollabSession[] = [];
const bindings: WordYjsCollaboration[] = [];
const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	for (const binding of bindings.splice(0)) binding.destroy();
	for (const session of sessions.splice(0)) session.destroy();
});

for (const atom of [
	{ text: '\n' },
	{ text: '', break: 'page' },
	{ text: '', noteReference: { kind: 'footnote', id: '1' } },
	{
		text: '',
		image: {
			relId: 'rId1',
			partName: 'word/media/a.png',
			contentType: 'image/png',
			widthPx: 10,
			heightPx: 20,
		},
	},
] satisfies TextRun[])
	for (const removeBold of [false, true])
		it(`merges independent concurrent ${atom.text ? 'hard-break' : Object.keys(atom)[1]} properties${removeBold ? ' including removal' : ''} and undoes only the local changes`, () => {
			let deliver = true;
			const hub = createMemoryHub({ filter: () => deliver });
			const peers = [0, 1].map((index) =>
				createCollabSession({
					roomId: 'properties',
					provider: transportProvider({ transport: hub.createTransport('properties') }),
					user: { name: String(index) },
					heartbeatMs: 0,
					teardown: false,
				}),
			);
			sessions.push(...peers);
			const initial = schema.node(
				'doc',
				null,
				schema.node('paragraph', null, [
					schema.text('L'),
					...runToInlineNodes(
						{ ...atom, bold: removeBold, fontFamily: 'Arial', fontSize: 12 },
						schema,
					),
					schema.text('R'),
				]),
			);
			for (const [index, session] of peers.entries()) {
				const binding = new WordYjsCollaboration(session, initial, {
					documentId: 'source',
					initializeIfEmpty: index === 0,
				});
				bindings.push(binding);
				const state = EditorState.create(binding.state(schema));
				const view = new EditorView(document.createElement('div'), { state });
				views.push(view);
				view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2, 3)));
			}
			const [a, b] = views;
			deliver = false;
			applyRunFormattingPatch({ bold: removeBold ? undefined : true, fontFamily: 'Georgia' })(
				a!.state,
				a!.dispatch,
				a,
			);
			applyRunFormattingPatch({ italic: true, fontSize: 18 })(b!.state, b!.dispatch, b);
			deliver = true;
			peers[0]!.resync();
			expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
			expect(inlineNodeRun(a!.state.doc.nodeAt(2)!)).toMatchObject({
				...atom,
				italic: true,
				fontFamily: 'Georgia',
				fontSize: 18,
			});
			expect(inlineNodeRun(a!.state.doc.nodeAt(2)!)!.bold).toBe(removeBold ? undefined : true);
			expect(bindings[0]!.undo()).toBe(true);
			expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
			expect(inlineNodeRun(b!.state.doc.nodeAt(2)!)).toMatchObject({
				...atom,
				bold: removeBold,
				italic: true,
				fontFamily: 'Arial',
				fontSize: 18,
			});
			expect(bindings[1]!.undo()).toBe(true);
			expect(a!.state.doc.eq(initial)).toBe(true);
			expect(b!.state.doc.eq(initial)).toBe(true);
		});
