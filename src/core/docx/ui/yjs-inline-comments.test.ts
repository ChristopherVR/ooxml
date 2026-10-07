// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, NodeSelection } from 'prosemirror-state';
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
import {
	hardBreakNodeSpec,
	pageBreakNodeSpec,
	noteReferenceNodeSpec,
	fieldMarkerNodeSpec,
} from './break-note-schema';
import { markSpecs } from './schema-marks';
import { runToInlineNodes, inlineNodeRun } from './run-adapter';
import { commentIdsAtSelection } from './comment-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0] },
		text: { group: 'inline' },
		image: imageNodeSpec,
		hardBreak: hardBreakNodeSpec,
		pageBreak: pageBreakNodeSpec,
		noteReference: noteReferenceNodeSpec,
		fieldMarker: fieldMarkerNodeSpec,
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
	{ text: '', fieldChar: 'begin' },
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
	it(`merges independent ${atom.text ? 'hard-break' : Object.keys(atom)[1]} comments and deletions with isolated author undo`, () => {
		let deliver = true;
		const hub = createMemoryHub({ filter: () => deliver });
		const peers = [0, 1].map((index) =>
			createCollabSession({
				roomId: 'inline-comments',
				provider: transportProvider({ transport: hub.createTransport('inline-comments') }),
				user: { name: String(index) },
				heartbeatMs: 0,
				teardown: false,
			}),
		);
		const errors: string[] = [];
		for (const peer of peers) peer.on('error', (error) => errors.push(String(error)));
		sessions.push(...peers);
		const initial = schema.node(
			'doc',
			null,
			schema.node('paragraph', null, [
				schema.text('L'),
				...runToInlineNodes({ ...atom, bold: true, commentIds: ['first', 'second'] }, schema),
				schema.text('R'),
			]),
		);
		for (const [index, session] of peers.entries()) {
			const binding = new WordYjsCollaboration(session, initial, {
				documentId: 'source',
				initializeIfEmpty: index === 0,
				initialComments: [
					{ id: 'first', author: 'Ada', text: 'First' },
					{ id: 'second', author: 'Bob', text: 'Second' },
				],
			});
			bindings.push(binding);
			const view = new EditorView(document.createElement('div'), {
				state: EditorState.create(binding.state(schema)),
			});
			views.push(view);
			view.dispatch(
				view.state.tr.setSelection(
					atom.image
						? NodeSelection.create(view.state.doc, 2)
						: TextSelection.create(view.state.doc, 2, 3),
				),
			);
		}
		const [a, b] = views;
		const ids = () =>
			views.map((view) => inlineNodeRun(view.state.doc.nodeAt(2)!)?.commentIds ?? []);
		expect(ids()).toEqual([
			['first', 'second'],
			['first', 'second'],
		]);
		deliver = false;
		expect(bindings[0]!.comments.add(a!, 'Ada', 'A', () => 'a')).not.toBeNull();
		expect(bindings[1]!.comments.add(b!, 'Bob', 'B', () => 'b')).not.toBeNull();
		deliver = true;
		peers[0]!.resync();
		expect(ids()).toEqual([
			['a', 'b', 'first', 'second'],
			['a', 'b', 'first', 'second'],
		]);
		expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
		expect(bindings[0]!.comments.all()).toHaveLength(4);
		expect(bindings[0]!.undo()).toBe(true);
		expect(ids()).toEqual([
			['b', 'first', 'second'],
			['b', 'first', 'second'],
		]);
		expect(bindings[0]!.comments.all().map((comment) => comment.id)).toEqual([
			'b',
			'first',
			'second',
		]);
		expect(bindings[0]!.redo()).toBe(true);
		deliver = false;
		expect(bindings[0]!.comments.delete(a!, 'first')).toBe(true);
		expect(bindings[1]!.comments.delete(b!, 'second')).toBe(true);
		deliver = true;
		peers[0]!.resync();
		expect(ids()).toEqual([
			['a', 'b'],
			['a', 'b'],
		]);
		expect(bindings[0]!.undo()).toBe(true);
		expect(ids()).toEqual([
			['a', 'b', 'first'],
			['a', 'b', 'first'],
		]);
		expect(commentIdsAtSelection(b!)).toEqual(['a', 'b', 'first']);
		expect(inlineNodeRun(b!.state.doc.nodeAt(2)!)?.bold).toBe(true);
		bindings[0]!.stopCapturing();
		a!.dispatch(a!.state.tr.insertText('Z', 3));
		expect(inlineNodeRun(b!.state.doc.nodeAt(2)!)?.commentIds).toEqual(['a', 'b', 'first']);
		a!.dispatch(a!.state.tr.insertText('Q', 2));
		expect(inlineNodeRun(b!.state.doc.nodeAt(3)!)?.commentIds).toEqual(['a', 'b', 'first']);
		a!.dispatch(a!.state.tr.insertText('!', 1));
		expect(inlineNodeRun(b!.state.doc.nodeAt(4)!)?.commentIds).toEqual(['a', 'b', 'first']);
		a!.dispatch(a!.state.tr.delete(4, 5));
		expect(errors).toEqual([]);
		expect(b!.state.doc.textContent).toBe('!LQZR');
		expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
		expect(commentIdsAtSelection(b!)).toEqual([]);
	});
