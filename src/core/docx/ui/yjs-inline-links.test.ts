// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { DOMParser, DOMSerializer, Schema } from 'prosemirror-model';
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
import { applyWordLink, removeWordLink, wordLinkRangeAt } from './link-commands';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0], parseDOM: [{ tag: 'p' }] },
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

it('retargets one contiguous text and picture link while preserving formatting and command probes', () => {
	const target = { href: 'https://example.com' };
	const picture: TextRun = {
		text: '',
		bold: true,
		link: target,
		image: {
			relId: 'rId1',
			partName: 'word/media/a.png',
			contentType: 'image/png',
			widthPx: 10,
			heightPx: 20,
		},
	};
	const doc = schema.node(
		'doc',
		null,
		schema.node('paragraph', null, [
			...runToInlineNodes({ text: 'L', italic: true, link: target }, schema),
			...runToInlineNodes(picture, schema),
			...runToInlineNodes({ text: 'R', link: target }, schema),
		]),
	);
	let state = EditorState.create({ doc, selection: TextSelection.create(doc, 2) });
	expect(wordLinkRangeAt(state, 2)).toMatchObject({ from: 1, to: 4 });
	const command = applyWordLink({ anchor: 'Destination' });
	expect(command(state)).toBe(true);
	expect(state.doc.eq(doc)).toBe(true);
	expect(command(state, undefined, { editable: false } as EditorView)).toBe(false);
	command(state, (tr) => {
		state = state.apply(tr);
	});
	for (const node of [state.doc.nodeAt(1)!, state.doc.nodeAt(2)!, state.doc.nodeAt(3)!])
		expect(inlineNodeRun(node)?.link).toEqual({ anchor: 'Destination' });
	expect(inlineNodeRun(state.doc.nodeAt(1)!)?.italic).toBe(true);
	expect(inlineNodeRun(state.doc.nodeAt(2)!)?.bold).toBe(true);
	removeWordLink(state, (tr) => {
		state = state.apply(tr);
	});
	expect(wordLinkRangeAt(state, 2)).toBeUndefined();
	expect(inlineNodeRun(state.doc.nodeAt(2)!)?.image).toEqual(picture.image);
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
	it(`preserves and edits ${atom.text ? 'hard-break' : Object.keys(atom)[1]} links across Yjs, copy and local undo`, () => {
		const source = {
			...atom,
			bold: true,
			link: { href: 'https://example.com', tooltip: 'Example' },
		};
		const initial = schema.node(
			'doc',
			null,
			schema.node('paragraph', null, runToInlineNodes(source, schema)),
		);
		const hub = createMemoryHub();
		const peers = [0, 1].map((index) =>
			createCollabSession({
				roomId: 'links',
				provider: transportProvider({ transport: hub.createTransport('links') }),
				user: { name: String(index) },
				heartbeatMs: 0,
				teardown: false,
			}),
		);
		sessions.push(...peers);
		for (const [index, session] of peers.entries()) {
			const binding = new WordYjsCollaboration(session, initial, {
				documentId: 'source',
				initializeIfEmpty: index === 0,
			});
			bindings.push(binding);
			views.push(
				new EditorView(document.createElement('div'), {
					state: EditorState.create(binding.state(schema)),
				}),
			);
		}
		const [a, b] = views;
		expect(inlineNodeRun(b!.state.doc.nodeAt(1)!)).toEqual(source);
		expect(wordLinkRangeAt(b!.state, 1)?.mark.attrs).toMatchObject(source.link);
		b!.dispatch(b!.state.tr.setSelection(TextSelection.create(b!.state.doc, 1, 2)));
		applyWordLink({ anchor: 'Destination', tooltip: 'Jump' })(b!.state, b!.dispatch, b);
		expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
		expect(inlineNodeRun(a!.state.doc.nodeAt(1)!)?.link).toEqual({
			anchor: 'Destination',
			tooltip: 'Jump',
		});
		const host = document.createElement('div');
		host.append(DOMSerializer.fromSchema(schema).serializeFragment(a!.state.doc.content));
		expect(inlineNodeRun(DOMParser.fromSchema(schema).parse(host).nodeAt(1)!)?.link).toEqual({
			anchor: 'Destination',
			tooltip: 'Jump',
		});
		expect(bindings[1]!.undo()).toBe(true);
		expect(inlineNodeRun(a!.state.doc.nodeAt(1)!)).toEqual(source);
		bindings[1]!.stopCapturing();
		b!.dispatch(b!.state.tr.setSelection(TextSelection.create(b!.state.doc, 1, 1)));
		removeWordLink(b!.state, b!.dispatch, b);
		expect(a!.state.doc.eq(b!.state.doc)).toBe(true);
		expect(inlineNodeRun(a!.state.doc.nodeAt(1)!)?.link).toBeUndefined();
		expect(inlineNodeRun(a!.state.doc.nodeAt(1)!)?.bold).toBe(true);
		expect(bindings[1]!.undo()).toBe(true);
		expect(a!.state.doc.eq(initial)).toBe(true);
	});
