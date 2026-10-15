import * as Y from 'yjs';
import {
	type ActivityItem,
	ACTIVITY_FILTERS,
	deriveActivity,
	describeActivity,
	filterActivity,
} from './activity';
import { createChatStore } from './chat';
import { createActivityReadState } from './activity-state';
import type { MentionTarget } from './mentions';

const ada: MentionTarget = { kind: 'user', id: 'ada', name: 'Ada' };
const team: MentionTarget = { kind: 'team', id: 'team', name: 'team' };

/** Two people on two documents that converge after every action, like two synced clients. */
function setup() {
	const docA = new Y.Doc();
	const docB = new Y.Doc();
	const sync = (): void => {
		Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA));
		Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB));
	};
	const synced = <T extends object>(store: T): T =>
		new Proxy(store, {
			get(target, key) {
				const value = (target as Record<string | symbol, unknown>)[key];
				if (typeof value !== 'function') return value;
				return (...args: unknown[]) => {
					sync();
					const result = (value as (...a: unknown[]) => unknown)(...args);
					sync();
					return result;
				};
			},
		});
	const me = synced(createChatStore(docA, { id: 'ada', name: 'Ada' }));
	const bob = synced(createChatStore(docB, { id: 'bob', name: 'Bob' }));
	const ops = me.createChannel({ name: 'Ops', id: 'ops' })!;
	const derive = (over: Partial<Parameters<typeof deriveActivity>[0]> = {}) =>
		deriveActivity({
			selfId: 'ada',
			channels: me.channels(),
			messagesOf: (id) => me.messages(id),
			reactions: me.reactionEvents(),
			pinsOf: (id) => me.pinned(id),
			...over,
		});
	return { me, bob, ops, derive };
}

describe('deriveActivity', () => {
	it('lists mentions of you, newest first, never your own or others mentions', () => {
		vi.useFakeTimers();
		try {
			const { me, bob, ops, derive } = setup();
			vi.setSystemTime(1000);
			bob.post(ops.id, { text: 'hi @Ada', mentions: [ada] });
			vi.setSystemTime(2000);
			bob.post(ops.id, { text: 'hello @team', mentions: [team] });
			vi.setSystemTime(3000);
			me.post(ops.id, { text: 'note to self @Ada', mentions: [ada] });
			bob.post(ops.id, { text: 'hey @Bob', mentions: [{ kind: 'user', id: 'bob', name: 'Bob' }] });
			const items = derive();
			expect(items.map((i) => [i.kind, i.preview])).toEqual([
				['mention', 'hello @team'],
				['mention', 'hi @Ada'],
			]);
			expect(items[0]).toMatchObject({ actorId: 'bob', actorName: 'Bob', channelName: 'Ops' });
		} finally {
			vi.useRealTimers();
		}
	});

	it('lists replies to your messages and to followed threads, with a mention taking priority', () => {
		const { me, bob, ops, derive } = setup();
		const mine = me.post(ops.id, { text: 'my post' })!;
		const theirs = bob.post(ops.id, { text: 'their post' })!;
		const reply = bob.post(ops.id, { text: 'answer', replyTo: mine.id })!;
		const followed = bob.post(ops.id, { text: 'followed reply', replyTo: theirs.id })!;
		const ping = bob.post(ops.id, { text: 'also @Ada', replyTo: mine.id, mentions: [ada] })!;
		bob.post(ops.id, { text: 'unrelated reply', replyTo: theirs.id });
		const none = derive();
		expect(none.filter((i) => i.kind === 'reply').map((i) => i.messageId)).toEqual([reply.id]);
		expect(none.find((i) => i.messageId === ping.id)?.kind).toBe('mention');
		expect(none.find((i) => i.messageId === ping.id)?.rootId).toBe(mine.id);
		const follows = derive({ isFollowed: (_c, root) => root === theirs.id });
		expect(follows.filter((i) => i.kind === 'reply').map((i) => i.messageId)).toContain(
			followed.id,
		);
	});

	it('lists reactions to your messages by others and pins by others', () => {
		const { me, bob, ops, derive } = setup();
		const mine = me.post(ops.id, { text: 'mine' })!;
		const theirs = bob.post(ops.id, { text: 'theirs' })!;
		bob.react(mine.id, '👍');
		bob.react(theirs.id, '🎉');
		me.react(mine.id, '❤️');
		bob.pin(ops.id, theirs.id);
		me.pin(ops.id, mine.id);
		const items = derive();
		expect(items.map((i) => i.kind).sort()).toEqual(['pin', 'reaction']);
		const reaction = items.find((i) => i.kind === 'reaction')!;
		expect(reaction).toMatchObject({ emoji: '👍', actorName: 'Bob', messageId: mine.id });
		expect(describeActivity(reaction)).toBe('Bob reacted 👍 to your message');
		expect(describeActivity(items.find((i) => i.kind === 'pin')!)).toBe(
			'Bob pinned a message in # Ops',
		);
	});

	it('skips deleted messages, archived channels and reactions from older clients', () => {
		const { me, bob, ops, derive } = setup();
		const mine = me.post(ops.id, { text: 'mine' })!;
		const gone = bob.post(ops.id, { text: 'hi @Ada', mentions: [ada] })!;
		bob.remove(ops.id, gone.id);
		expect(derive()).toEqual([]);
		bob.react(mine.id, '👍');
		expect(
			derive({ reactions: [{ messageId: mine.id, emoji: '👍', userId: 'bob', ts: 0 }] }),
		).toEqual([]);
		me.archiveChannel(ops.id);
		expect(derive()).toEqual([]);
	});

	it('applies read state and filters like the Teams feed', () => {
		const { me, bob, ops, derive } = setup();
		const mine = me.post(ops.id, { text: 'mine' })!;
		const first = bob.post(ops.id, { text: 'hi @Ada', mentions: [ada] })!;
		bob.post(ops.id, { text: 'reply', replyTo: mine.id });
		bob.react(mine.id, '👍');
		const storage = new Map<string, string>();
		const state = createActivityReadState(
			{ getItem: (k) => storage.get(k) ?? null, setItem: (k, v) => void storage.set(k, v) },
			'k',
			() => 0,
		);
		expect(state.mark(`mention:${first.id}`, first.ts, true)).toBe(true);
		const items = derive({ isRead: state.isRead });
		const count = (f: (typeof ACTIVITY_FILTERS)[number]): number =>
			filterActivity(items as ActivityItem[], f).length;
		expect([
			count('all'),
			count('unread'),
			count('mentions'),
			count('replies'),
			count('reactions'),
		]).toEqual([3, 2, 1, 1, 1]);
	});
});
