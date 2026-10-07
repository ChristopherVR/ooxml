import { afterEach, describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from '../../collab/index';
import { WordYjsCollaboration } from './yjs-collaboration';

const schema = new Schema({
	nodes: {
		doc: {
			content: 'paragraph+',
			attrs: { sections: { default: null }, sectionParts: { default: null } },
		},
		paragraph: { content: 'text*' },
		text: {},
	},
});
const sessions: CollabSession[] = [];
const bindings: WordYjsCollaboration[] = [];
afterEach(() => {
	for (const binding of bindings.splice(0)) binding.destroy();
	for (const session of sessions.splice(0)) session.destroy();
});
function pair(autoConnect = true) {
	const hub = createMemoryHub();
	return [0, 1].map((index) => {
		const session = createCollabSession({
			roomId: 'word',
			provider: transportProvider({ transport: hub.createTransport('word') }),
			user: { name: String(index) },
			autoConnect,
			heartbeatMs: 0,
			teardown: false,
		});
		sessions.push(session);
		return session;
	});
}
function bind(
	session: CollabSession,
	initial: ReturnType<Schema['node']>,
	initializeIfEmpty = false,
) {
	const binding = new WordYjsCollaboration(session, initial, {
		documentId: 'package-id',
		initializeIfEmpty,
	});
	bindings.push(binding);
	return binding;
}

describe('Word room bootstrap', () => {
	it('preserves sections and header/footer data instead of adopting the joining snapshot', () => {
		const [a, b] = pair();
		const initial = schema.node(
			'doc',
			{ sections: '[{"breakType":"continuous"}]', sectionParts: '[{"header":"retained"}]' },
			[schema.node('paragraph', null, schema.text('creator'))],
		);
		bind(a!, initial, true);
		const other = schema.node('doc', null, [
			schema.node('paragraph', null, schema.text('stale snapshot')),
		]);
		const joined = bind(b!, other).state(schema);
		expect(joined.doc.toJSON()).toEqual(initial.toJSON());
		expect(b!.doc.getMap('docx:identity').get('documentId')).toBe('package-id');
	});
	it('requires actual synchronization before seeding a room', () => {
		const [a] = pair(false);
		const initial = schema.node('doc', null, [schema.node('paragraph')]);
		expect(() => bind(a!, initial, true)).toThrow(/initial provider synchronization/);
		expect(a!.doc.getXmlFragment('docx:body').length).toBe(0);
	});
	it('creates a shared empty text container without adding initial content to local undo', () => {
		const [a, b] = pair();
		const initial = schema.node('doc', null, [schema.node('paragraph')]);
		const binding = bind(a!, initial, true);
		bind(b!, initial);
		expect(binding.fragment.get(0).toString()).toBe('<paragraph></paragraph>');
		expect(binding.canUndo()).toBe(false);
		binding.destroy();
		binding.destroy();
		expect(() => binding.state(schema)).toThrow(/destroyed/);
		expect(a!.status).toBe('connected');
	});
});
