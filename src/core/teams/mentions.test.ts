import * as Y from 'yjs';
import { createChatStore } from './chat';
import {
	GROUP_MENTIONS,
	type MentionTarget,
	filterMentionTargets,
	insertMention,
	mentionQueryAt,
	mentionSegments,
	mentionTargetsFor,
	mentionsUser,
	resolveMentions,
	sanitizeMentions,
} from './mentions';

const ada: MentionTarget = { kind: 'user', id: 'ada', name: 'Ada' };
const adaL: MentionTarget = { kind: 'user', id: 'ada2', name: 'Ada Lovelace' };
const bob: MentionTarget = { kind: 'user', id: 'bob', name: 'Bob' };
const [channel, team] = GROUP_MENTIONS as [MentionTarget, MentionTarget];

describe('resolveMentions', () => {
	it('finds mentions, the longest name winning, case-insensitively', () => {
		const text = 'hi @Ada Lovelace and @bob, also @ada';
		const found = resolveMentions(text, [ada, adaL, bob]);
		expect(found.map((m) => [m.id, text.slice(m.start, m.end)])).toEqual([
			['ada2', '@Ada Lovelace'],
			['bob', '@bob'],
			['ada', '@ada'],
		]);
	});

	it('needs word boundaries, so addresses and longer words are not mentions', () => {
		expect(resolveMentions('mail ada@Ada.com or @Adam', [ada])).toEqual([]);
		expect(resolveMentions('@Ada.', [ada])).toHaveLength(1);
		expect(resolveMentions('(@Ada)', [ada])).toHaveLength(1);
	});

	it('resolves @channel and @team and ignores malformed targets', () => {
		const found = resolveMentions('@channel heads up, @team too', [
			channel,
			team,
			{ kind: 'channel', id: 'x', name: 'everyone' } as MentionTarget,
			{ kind: 'user', id: '', name: 'Nobody' },
		]);
		expect(found.map((m) => m.kind)).toEqual(['channel', 'team']);
	});
});

describe('sanitizeMentions', () => {
	it('keeps only ranges that still sit on their name', () => {
		const text = 'hello @Ada!';
		const ok = { ...ada, start: 6, end: 10 };
		expect(sanitizeMentions([ok], text)).toEqual([ok]);
		expect(sanitizeMentions([{ ...ok, start: 5, end: 9 }], text)).toEqual([]);
		expect(sanitizeMentions([{ ...ok, end: 99 }], text)).toEqual([]);
		expect(sanitizeMentions([ok, ok], text)).toEqual([ok]);
		expect(sanitizeMentions('nope', text)).toEqual([]);
		expect(sanitizeMentions([{ kind: 'user', id: 'a', name: '', start: 0, end: 1 }], '@')).toEqual(
			[],
		);
	});
});

describe('mentionSegments and mentionsUser', () => {
	it('splits text around mentions', () => {
		const text = 'a @Ada b';
		const mentions = resolveMentions(text, [ada]);
		expect(mentionSegments(text, mentions).map((s) => [s.text, !!s.mention])).toEqual([
			['a ', false],
			['@Ada', true],
			[' b', false],
		]);
		expect(mentionSegments('plain')).toEqual([{ text: 'plain' }]);
	});

	it('treats @channel and @team as addressing everyone', () => {
		expect(mentionsUser({ mentions: resolveMentions('@Ada', [ada]) }, 'ada')).toBe(true);
		expect(mentionsUser({ mentions: resolveMentions('@Ada', [ada]) }, 'bob')).toBe(false);
		expect(mentionsUser({ mentions: resolveMentions('@team', [team]) }, 'bob')).toBe(true);
		expect(mentionsUser({}, 'bob')).toBe(false);
	});
});

describe('the picker rules', () => {
	it('finds the open query before the caret', () => {
		expect(mentionQueryAt('hi @Ad', 6)).toEqual({ start: 3, query: 'Ad' });
		expect(mentionQueryAt('hi @', 4)).toEqual({ start: 3, query: '' });
		expect(mentionQueryAt('hi @Ada Lo', 10)).toEqual({ start: 3, query: 'Ada Lo' });
		expect(mentionQueryAt('mail a@b', 8)).toBeNull();
		expect(mentionQueryAt('hi @Ada\nnext', 12)).toBeNull();
		expect(mentionQueryAt('no at', 5)).toBeNull();
		expect(mentionQueryAt('hi @Ada ', 8)).toBeNull();
	});

	it('filters by prefix, then word prefix, then substring', () => {
		const all = [bob, adaL, ada, channel];
		expect(filterMentionTargets(all, '').map((t) => t.name)).toEqual([
			'Bob',
			'Ada Lovelace',
			'Ada',
			'channel',
		]);
		expect(filterMentionTargets(all, 'ad').map((t) => t.name)).toEqual(['Ada Lovelace', 'Ada']);
		expect(filterMentionTargets(all, 'love').map((t) => t.name)).toEqual(['Ada Lovelace']);
		expect(filterMentionTargets(all, 'zz')).toEqual([]);
		expect(filterMentionTargets(all, '', 2)).toHaveLength(2);
	});

	it('inserts the token and moves the caret past it', () => {
		const query = mentionQueryAt('hi @Ad there', 6)!;
		expect(insertMention('hi @Ad there', query, 6, ada)).toEqual({
			text: 'hi @Ada there',
			caret: 8,
		});
		const end = mentionQueryAt('hi @', 4)!;
		expect(insertMention('hi @', end, 4, ada)).toEqual({ text: 'hi @Ada ', caret: 8 });
	});

	it('offers people who wrote or are online, never yourself, then the groups', () => {
		const targets = mentionTargetsFor(
			[
				{ authorId: 'bob', authorName: 'Bob' },
				{ authorId: 'me', authorName: 'Me' },
			],
			[
				{ id: 'ada', name: 'Ada' },
				{ id: 'bob', name: 'Bob' },
			],
			'me',
		);
		expect(targets.map((t) => t.name)).toEqual(['Ada', 'Bob', 'channel', 'team']);
	});
});

describe('mentions in the chat store', () => {
	const exchange = (a: Y.Doc, b: Y.Doc): void => {
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
	};

	it('stores structural mentions, syncs them and keeps plain text readable', () => {
		const docA = new Y.Doc();
		const docB = new Y.Doc();
		const a = createChatStore(docA, { id: 'bob', name: 'Bob' });
		const b = createChatStore(docB, { id: 'ada', name: 'Ada' });
		const c = a.createChannel({ name: 'Ops', id: 'ops' })!;
		const posted = a.post(c.id, { text: '  hello @Ada  ', mentions: [ada] })!;
		expect(posted.text).toBe('hello @Ada');
		expect(posted.mentions).toEqual([{ ...ada, start: 6, end: 10 }]);
		exchange(docA, docB);
		expect(b.messages(c.id)[0]?.mentions?.[0]).toMatchObject({ id: 'ada', start: 6, end: 10 });
	});

	it('reads messages without mentions (older clients) as plain text', () => {
		const doc = new Y.Doc();
		const store = createChatStore(doc, { id: 'bob', name: 'Bob' });
		const c = store.createChannel({ name: 'Ops', id: 'ops' })!;
		const posted = store.post(c.id, { text: 'hi @Ada' })!;
		expect(posted.mentions).toBeUndefined();
		// A hostile row pointing at the wrong text is dropped on read.
		const row = doc.getMap<Y.Map<unknown>>('teams:messages:ops').get(posted.id)!;
		row.set('mentions', [{ ...ada, start: 0, end: 3 }]);
		expect(store.messages(c.id)[0]?.mentions).toBeUndefined();
	});

	it('keeps a mention when the message is edited around it, and drops it when removed', () => {
		const store = createChatStore(new Y.Doc(), { id: 'bob', name: 'Bob' });
		const c = store.createChannel({ name: 'Ops', id: 'ops' })!;
		const m = store.post(c.id, { text: 'hi @Ada', mentions: [ada] })!;
		expect(store.edit(c.id, m.id, 'well, hi @Ada, welcome')).toBe(true);
		let edited = store.messages(c.id)[0]!;
		expect(edited.text).toBe('well, hi @Ada, welcome');
		expect(edited.mentions).toEqual([{ ...ada, start: 9, end: 13 }]);
		expect(store.edit(c.id, m.id, 'well, welcome', [bob])).toBe(true);
		edited = store.messages(c.id)[0]!;
		expect(edited.mentions).toBeUndefined();
	});

	it('timestamps reactions for the activity feed', () => {
		const store = createChatStore(new Y.Doc(), { id: 'bob', name: 'Bob' });
		const c = store.createChannel({ name: 'Ops', id: 'ops' })!;
		const m = store.post(c.id, { text: 'x' })!;
		store.react(m.id, '👍');
		const [event] = store.reactionEvents();
		expect(event).toMatchObject({ messageId: m.id, emoji: '👍', userId: 'bob' });
		expect(event!.ts).toBeGreaterThan(0);
		expect(store.messages(c.id)[0]?.reactions['👍']).toEqual(['bob']);
	});
});
