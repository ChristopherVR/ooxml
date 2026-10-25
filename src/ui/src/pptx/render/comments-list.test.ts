/* oxlint-disable eslint/one-var -- pervasive pre-existing pattern in this file
   (each spec sets up a couple of independent `const`s); merging them isn't a
   style choice here. */
import type { PptxComment, PptxCommentMention } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	addCommentToList,
	editCommentInList,
	removeCommentFromList,
	replyToCommentInList,
	toggleCommentResolvedInList,
} from './comments-list';

const parent = (overrides: Partial<PptxComment> = {}): PptxComment => ({
	id: 'parent-1',
	text: 'Top-level comment',
	author: 'Alice',
	resolved: false,
	...overrides,
});

describe('replyToCommentInList', () => {
	it('nests the reply under the parent with threading metadata', () => {
		const next = replyToCommentInList([parent()], 'parent-1', '  A reply  ', 'Bob');
		expect(next).not.toBeNull();
		const replies = next![0].replies;
		expect(replies).toHaveLength(1);
		expect(replies![0]).toMatchObject({
			text: 'A reply',
			author: 'Bob',
			threadId: 'parent-1',
			parentId: 'parent-1',
		});
		expect(replies![0].id).toBeTruthy();
		expect(replies![0].createdAt).toBeTruthy();
	});

	it('appends to existing replies without mutating the input', () => {
		const existingReply: PptxComment = { id: 'r1', text: 'First', parentId: 'parent-1' };
		const input = [parent({ replies: [existingReply] })];
		const next = replyToCommentInList(input, 'parent-1', 'Second', 'Bob');
		expect(next![0].replies).toHaveLength(2);
		expect(next![0].replies![0]).toBe(existingReply);
		expect(input[0].replies).toHaveLength(1);
	});

	it('inherits the parent elementId anchor', () => {
		const next = replyToCommentInList([parent({ elementId: 'shape-9' })], 'parent-1', 'Hi', 'Bob');
		expect(next![0].replies![0].elementId).toBe('shape-9');
	});

	it('returns null for blank text or an unknown parent', () => {
		expect(replyToCommentInList([parent()], 'parent-1', '   ', 'Bob')).toBeNull();
		expect(replyToCommentInList([parent()], 'missing', 'Hi', 'Bob')).toBeNull();
	});
});

describe('addCommentToList', () => {
	it('stamps the elementId anchor when provided', () => {
		const next = addCommentToList([], 'Anchored', 'Alice', undefined, undefined, 'shape-3');
		expect(next![0]).toMatchObject({ text: 'Anchored', author: 'Alice', elementId: 'shape-3' });
	});

	it('omits elementId when not provided', () => {
		const next = addCommentToList([], 'Unanchored', 'Alice');
		expect(next![0]).not.toHaveProperty('elementId');
	});
});

describe('nested-reply tree traversal (removeCommentFromList / toggleCommentResolvedInList / editCommentInList)', () => {
	const withNestedReply = (): PptxComment[] => [
		parent({
			replies: [
				{ id: 'r1', text: 'Reply one', parentId: 'parent-1', resolved: false },
				{ id: 'r2', text: 'Reply two', parentId: 'parent-1', resolved: false },
			],
		}),
	];

	it('removeCommentFromList drops a nested reply without touching its siblings', () => {
		const next = removeCommentFromList(withNestedReply(), 'r1');
		expect(next).not.toBeNull();
		expect(next![0].replies?.map((r) => r.id)).toStrictEqual(['r2']);
	});

	it('removeCommentFromList returns null when the id is not found anywhere', () => {
		expect(removeCommentFromList(withNestedReply(), 'missing')).toBeNull();
	});

	it('toggleCommentResolvedInList flips resolved on a nested reply, leaving the parent untouched', () => {
		const next = toggleCommentResolvedInList(withNestedReply(), 'r1');
		expect(next).not.toBeNull();
		expect(next![0].resolved).toBeFalsy();
		expect(next![0].replies?.[0].resolved).toBeTruthy();
		expect(next![0].replies?.[1].resolved).toBeFalsy();
	});

	it('editCommentInList rewrites a nested reply, leaving the parent text untouched', () => {
		const next = editCommentInList(withNestedReply(), 'r1', 'Edited reply');
		expect(next).not.toBeNull();
		expect(next![0].text).toBe('Top-level comment');
		expect(next![0].replies?.[0].text).toBe('Edited reply');
	});

	it('editCommentInList returns null for blank text or an unknown id', () => {
		expect(editCommentInList(withNestedReply(), 'r1', '   ')).toBeNull();
		expect(editCommentInList(withNestedReply(), 'missing', 'text')).toBeNull();
	});

	it('does not mutate the input tree', () => {
		const input = withNestedReply();
		removeCommentFromList(input, 'r1');
		toggleCommentResolvedInList(input, 'r1');
		editCommentInList(input, 'r1', 'changed');
		expect(input[0].replies).toHaveLength(2);
		expect(input[0].replies?.[0].text).toBe('Reply one');
		expect(input[0].replies?.[0].resolved).toBeFalsy();
	});
});

describe('mentions on new, replied and edited comments', () => {
	const GRACE = '{2CB2E9D0-D392-EB21-5D46-FBA34C1295E6}';
	const grace = (
		startIndex: number,
		extra: Partial<PptxCommentMention> = {},
	): PptxCommentMention => ({
		id: '{M1}',
		personId: GRACE,
		authorName: 'Grace',
		startIndex,
		length: 6,
		...extra,
	});
	const span = (comment: PptxComment | undefined) =>
		comment?.mentions?.map((m) => comment.text.slice(m.startIndex, m.startIndex + m.length));

	it('stores the spans of a new comment, moved with the trimmed text', () => {
		const next = addCommentToList([], '  @Grace hi ', 'Alice', undefined, undefined, undefined, [
			grace(2),
		]);
		expect(next![0]).toMatchObject({ text: '@Grace hi', mentions: [grace(0)] });
		expect(span(next![0])).toStrictEqual(['@Grace']);
		expect(
			addCommentToList([], 'Plain', 'Alice', undefined, undefined, undefined, [])![0],
		).not.toHaveProperty('mentions');
	});

	it('stores the spans of a reply, moved with the trimmed text', () => {
		const next = replyToCommentInList([parent()], 'parent-1', ' Ask @Grace', 'Bob', [grace(5)]);
		const reply = next![0].replies![0];
		expect(reply.mentions).toStrictEqual([grace(4)]);
		expect(span(reply)).toStrictEqual(['@Grace']);
	});

	it('keeps an edited comment spans on their text, or takes the ones given', () => {
		const comments = [parent({ text: 'Ask @Grace today', mentions: [grace(4)] })];
		const edited = editCommentInList(comments, 'parent-1', 'Please ask @Grace today');
		expect(edited![0].mentions).toStrictEqual([grace(11)]);
		expect(span(edited![0])).toStrictEqual(['@Grace']);
		// Edits on both sides of a mention keep it on its text.
		const around = editCommentInList(comments, 'parent-1', 'Please ask @Grace tomorrow');
		expect(span(around![0])).toStrictEqual(['@Grace']);
		// An edit inside the mention drops it, and the field goes with the last one.
		expect(editCommentInList(comments, 'parent-1', 'Ask @Grce today')![0]).not.toHaveProperty(
			'mentions',
		);
		const given = editCommentInList(comments, 'parent-1', ' @Grace, ask @Grace', [
			grace(1),
			grace(13, { id: '{M2}' }),
		]);
		expect(given![0].mentions).toStrictEqual([grace(0), grace(12, { id: '{M2}' })]);
	});
});
