import * as Y from 'yjs';
import { createChatStore } from './chat';
import { MAX_PINNED_MESSAGES } from './pins';

const exchange = (a: Y.Doc, b: Y.Doc): void => {
	Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
	Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
};

function setup() {
	const docA = new Y.Doc();
	const docB = new Y.Doc();
	const ada = createChatStore(docA, { id: 'ada', name: 'Ada' });
	const bob = createChatStore(docB, { id: 'bob', name: 'Bob' });
	const channel = ada.createChannel({ name: 'Ops', id: 'ops' })!;
	const first = ada.post(channel.id, { text: 'first' })!;
	const second = ada.post(channel.id, { text: 'second' })!;
	exchange(docA, docB);
	return { docA, docB, ada, bob, channel, first, second };
}

describe('pinned messages', () => {
	it('pins and unpins, listing the newest pin first', () => {
		const { ada, channel, first, second } = setup();
		vi.useFakeTimers();
		try {
			vi.setSystemTime(1000);
			expect(ada.pin(channel.id, first.id)).toBe(true);
			vi.setSystemTime(2000);
			expect(ada.pin(channel.id, second.id)).toBe(true);
		} finally {
			vi.useRealTimers();
		}
		expect(ada.pinned(channel.id).map((p) => p.message.id)).toEqual([second.id, first.id]);
		expect(ada.pinned(channel.id)[0]).toMatchObject({ pinnedBy: 'ada', pinnedByName: 'Ada' });
		expect(ada.pin(channel.id, first.id)).toBe(false);
		expect(ada.unpin(channel.id, first.id)).toBe(true);
		expect(ada.unpin(channel.id, first.id)).toBe(false);
		expect(ada.pinned(channel.id).map((p) => p.message.id)).toEqual([second.id]);
	});

	it('syncs between peers, and any member can unpin', () => {
		const { docA, docB, ada, bob, channel, first } = setup();
		ada.pin(channel.id, first.id);
		exchange(docA, docB);
		expect(bob.pinned(channel.id).map((p) => p.message.text)).toEqual(['first']);
		expect(bob.unpin(channel.id, first.id)).toBe(true);
		exchange(docA, docB);
		expect(ada.pinned(channel.id)).toEqual([]);
	});

	it('converges when two peers pin the same message at once', () => {
		const { docA, docB, ada, bob, channel, first } = setup();
		ada.pin(channel.id, first.id);
		bob.pin(channel.id, first.id);
		exchange(docA, docB);
		expect(ada.pinned(channel.id)).toHaveLength(1);
		expect(bob.pinned(channel.id)).toHaveLength(1);
	});

	it('hides pins of deleted messages and rejects unknown, deleted or archived targets', () => {
		const { ada, channel, first, second } = setup();
		expect(ada.pin(channel.id, 'missing')).toBe(false);
		expect(ada.pin('nope', first.id)).toBe(false);
		ada.pin(channel.id, first.id);
		ada.remove(channel.id, first.id);
		expect(ada.pinned(channel.id)).toEqual([]);
		expect(ada.pin(channel.id, first.id)).toBe(false);
		ada.archiveChannel(channel.id);
		expect(ada.pin(channel.id, second.id)).toBe(false);
	});

	it('ignores malformed rows written by a peer', () => {
		const { docA, ada, channel, first } = setup();
		const rows = docA.getMap<unknown>(`teams:pins:${channel.id}`);
		rows.set(first.id, { by: 42, at: 'x' });
		rows.set('bad id', { by: 'bob' });
		expect(ada.pinned(channel.id)).toEqual([]);
	});

	it('caps the pinned list', () => {
		const { ada, channel } = setup();
		for (let i = 0; i < MAX_PINNED_MESSAGES + 1; i++) {
			const m = ada.post(channel.id, { text: `m${i}` })!;
			expect(ada.pin(channel.id, m.id)).toBe(i < MAX_PINNED_MESSAGES);
		}
		expect(ada.pinned(channel.id)).toHaveLength(MAX_PINNED_MESSAGES);
	});
});
