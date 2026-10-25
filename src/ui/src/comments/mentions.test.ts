import { describe, expect, it } from 'vitest';

import {
	commentSegments,
	followTextEdit,
	insertMention,
	matchMentionCandidates,
	mentionQueryAt,
	type OfficeCommentMention,
	type OfficeMentionCandidate,
} from './mentions';

const ADA: OfficeMentionCandidate = { id: 'p1', name: 'Ada Lovelace', initials: 'AL' };
const GRACE: OfficeMentionCandidate = { id: 'p2', name: 'Grace Hopper' };
const ALAN: OfficeMentionCandidate = { id: 'p3', name: 'Alan Turing' };

describe('commentSegments', () => {
	it('splits the text into plain runs and mentions', () => {
		const mentions: OfficeCommentMention[] = [
			{ start: 4, length: 13, id: 'p1', name: 'Ada Lovelace' },
		];
		expect(commentSegments('Ask @Ada Lovelace now', mentions)).toStrictEqual([
			{ kind: 'text', text: 'Ask ' },
			{ kind: 'mention', text: '@Ada Lovelace', id: 'p1', name: 'Ada Lovelace' },
			{ kind: 'text', text: ' now' },
		]);
		expect(commentSegments('', mentions)).toStrictEqual([]);
		expect(commentSegments('plain')).toStrictEqual([{ kind: 'text', text: 'plain' }]);
	});

	it('ignores spans out of range, empty or overlapping an earlier one', () => {
		const spans: OfficeCommentMention[] = [
			{ start: 0, length: 2 },
			{ start: 1, length: 2 },
			{ start: 3, length: 0 },
			{ start: 4, length: 9 },
		];
		expect(commentSegments('abcdef', spans)).toStrictEqual([
			{ kind: 'mention', text: 'ab' },
			{ kind: 'text', text: 'cdef' },
		]);
	});
});

describe('mentionQueryAt', () => {
	it('finds the @ token at the caret, not one inside a word or after a space', () => {
		expect(mentionQueryAt('Hi @Gr', 6)).toStrictEqual({ query: 'gr', start: 3 });
		expect(mentionQueryAt('@', 1)).toStrictEqual({ query: '', start: 0 });
		expect(mentionQueryAt('mail ada@example.com', 20)).toBeNull();
		expect(mentionQueryAt('Hi @Gr ace', 10)).toBeNull();
		expect(mentionQueryAt('no token', 3)).toBeNull();
	});
});

describe('matchMentionCandidates', () => {
	it('ranks name prefixes first, then names or initials containing the query', () => {
		const all = [GRACE, ALAN, ADA];
		expect(matchMentionCandidates(all, '').map((c) => c.id)).toStrictEqual(['p1', 'p3', 'p2']);
		expect(matchMentionCandidates(all, 'a').map((c) => c.id)).toStrictEqual(['p1', 'p3', 'p2']);
		expect(matchMentionCandidates(all, 'al').map((c) => c.id)).toStrictEqual(['p3', 'p1']);
		expect(matchMentionCandidates(all, 'hop').map((c) => c.id)).toStrictEqual(['p2']);
		expect(matchMentionCandidates(all, 'zz')).toStrictEqual([]);
		expect(matchMentionCandidates(all, '', 2)).toHaveLength(2);
	});
});

describe('insertMention', () => {
	it('replaces the @ token with the name and records the span, @ included', () => {
		expect(insertMention('Ask @Ad', [], 7, ADA)).toStrictEqual({
			text: 'Ask @Ada Lovelace ',
			mentions: [{ start: 4, length: 13, id: 'p1', name: 'Ada Lovelace' }],
			caret: 18,
		});
	});

	it('shifts the mentions after it', () => {
		const later: OfficeCommentMention = { start: 6, length: 6, id: 'p2', name: 'Grace Hopper' };
		const result = insertMention('@G and @Grace', [{ ...later, start: 7 }], 2, GRACE);
		expect(result.text).toBe('@Grace Hopper  and @Grace');
		expect(result.mentions.map((m) => [m.start, m.length])).toStrictEqual([
			[0, 13],
			[19, 6],
		]);
	});
});

describe('followTextEdit', () => {
	const spans: OfficeCommentMention[] = [
		{ start: 0, length: 4, id: 'p1' },
		{ start: 9, length: 4, id: 'p2' },
	];

	it('keeps spans before the edit, shifts those after it and drops the ones it touches', () => {
		// "@Ada and @Bob" -> "@Ada, and @Bob": typed after the first span.
		expect(followTextEdit('@Ada and @Bob', '@Ada, and @Bob', spans)).toStrictEqual([
			{ start: 0, length: 4, id: 'p1' },
			{ start: 10, length: 4, id: 'p2' },
		]);
		// A letter deleted inside the first mention drops it.
		expect(followTextEdit('@Ada and @Bob', '@Aa and @Bob', spans)).toStrictEqual([
			{ start: 8, length: 4, id: 'p2' },
		]);
		expect(followTextEdit('@Ada and @Bob', '@Ada and @Bob', spans)).toBe(spans);
	});
});
