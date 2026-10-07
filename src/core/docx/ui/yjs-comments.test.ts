import { afterEach, describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from '../../collab/index.js';
import { WordYjsCollaboration } from './yjs-collaboration.js';

const schema = new Schema({
	nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*' }, text: {} },
});
const initial = schema.node('doc', null, [schema.node('paragraph', null, schema.text('Text'))]);
const sessions: CollabSession[] = [];
const bindings: WordYjsCollaboration[] = [];
afterEach(() => {
	for (const binding of bindings.splice(0)) binding.destroy();
	for (const session of sessions.splice(0)) session.destroy();
});
function pair() {
	let deliver = true;
	const hub = createMemoryHub({ filter: () => deliver });
	const peers = ['Ada', 'Grace'].map((name) => {
		const session = createCollabSession({
			roomId: 'word-comments',
			user: { name },
			provider: transportProvider({ transport: hub.createTransport('word-comments') }),
			heartbeatMs: 0,
			teardown: false,
		});
		sessions.push(session);
		return session;
	});
	const word = peers.map((session, index) => {
		const binding = new WordYjsCollaboration(session, initial, {
			documentId: 'source',
			initializeIfEmpty: index === 0,
			initialComments: [{ id: 'root', author: 'Ada', text: 'Review' }],
		});
		bindings.push(binding);
		return binding;
	});
	return {
		a: word[0]!,
		b: word[1]!,
		partition: () => {
			deliver = false;
		},
		sync: () => {
			deliver = true;
			peers[0]!.resync();
		},
	};
}
describe('DOM-free Word Yjs thread records', () => {
	it('merges independent replies and resolution without replacing the thread snapshot', () => {
		const { a, b, partition, sync } = pair();
		partition();
		a.comments.reply('root', 'Ada', 'A reply', () => 'a');
		b.comments.reply('root', 'Grace', 'B reply', () => 'b');
		b.comments.resolve('root', true);
		sync();
		expect(a.comments.all()).toEqual(b.comments.all());
		expect(a.comments.all()).toHaveLength(3);
		a.undo();
		expect(b.comments.all()).toEqual([
			{ id: 'b', parentId: 'root', author: 'Grace', text: 'B reply' },
			{ id: 'root', author: 'Ada', text: 'Review', resolved: true },
		]);
		const snapshot = a.comments.all();
		snapshot[0]!.text = 'External mutation';
		expect(a.comments.all()[0]!.text).toBe('B reply');
	});
	it('does not enable thread writes for a legacy room without the anchor capability', () => {
		const { a, b } = pair();
		a.session.doc.getMap('docx:identity').delete('comments');
		const legacy = new WordYjsCollaboration(b.session, initial, { documentId: 'source' });
		bindings.push(legacy);
		expect(legacy.sharedComments).toBe(false);
		expect(legacy.comments.reply('root', 'Grace', 'Unsupported', () => 'reply')).toBe(false);
		expect(legacy.comments.resolve('root', true)).toBe(false);
	});
});
