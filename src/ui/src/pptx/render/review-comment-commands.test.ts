// @vitest-environment happy-dom
import type { PptxComment } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import {
	adjacentComment,
	clearAllComments,
	clearSlideComments,
	findCommentLocation,
	listCommentLocations,
	removeCommentFromSlides,
	revealComment,
	trackActiveComment,
} from './comment-navigation';
import type { ReviewCommentsHost, SlideCommentsChange } from './review-comment-commands';
import { commentCounts, runReviewCommentCommand } from './review-comment-commands';

const c = (id: string, replies?: PptxComment[]): PptxComment => ({
	id,
	text: id,
	...(replies ? { replies } : {}),
});

const deck = () => [
	{ comments: [c('a'), c('b', [c('b1')])] },
	{ comments: [] as PptxComment[] },
	{ comments: [c('d')] },
	{},
];

describe('comment navigation', () => {
	it('lists top-level comments in slide order', () => {
		expect(listCommentLocations(deck()).map((l) => `${l.slideIndex}:${l.commentId}`)).toStrictEqual(
			['0:a', '0:b', '2:d'],
		);
	});

	it('resolves a reply to its thread', () => {
		expect(findCommentLocation(deck(), 'b1')).toStrictEqual({ slideIndex: 0, commentId: 'b' });
		expect(findCommentLocation(deck(), 'zzz')).toBeNull();
	});

	it('steps next and previous across slides and wraps', () => {
		const slides = deck();
		const at = (slideIndex: number, commentId: string) => ({ slideIndex, commentId });
		expect(adjacentComment(slides, at(0, 'b'), 0, 1)).toStrictEqual(at(2, 'd'));
		expect(adjacentComment(slides, at(2, 'd'), 2, 1)).toStrictEqual(at(0, 'a'));
		expect(adjacentComment(slides, at(0, 'a'), 0, -1)).toStrictEqual(at(2, 'd'));
		// No current comment: start from the current slide.
		expect(adjacentComment(slides, null, 1, 1)).toStrictEqual(at(2, 'd'));
		expect(adjacentComment(slides, null, 1, -1)).toStrictEqual(at(0, 'b'));
		expect(adjacentComment([{}, {}], null, 0, 1)).toBeNull();
	});

	it('removes one comment, a slide of comments, or all', () => {
		const slides = deck();
		expect(removeCommentFromSlides(slides, 'a')?.[0]?.comments?.map((x) => x.id)).toStrictEqual([
			'b',
		]);
		expect(removeCommentFromSlides(slides, 'b1')?.[0]?.comments?.[1]?.replies).toStrictEqual([]);
		expect(removeCommentFromSlides(slides, 'nope')).toBeNull();
		expect(clearSlideComments(slides, 0)?.[0]?.comments).toStrictEqual([]);
		expect(clearSlideComments(slides, 1)).toBeNull();
		expect(clearAllComments(slides)?.every((s) => !s.comments?.length)).toBe(true);
		expect(clearAllComments([{}, { comments: [] }])).toBeNull();
	});

	it('counts the slide and the deck', () => {
		expect(commentCounts(deck(), 0)).toStrictEqual({ slide: 2, total: 3 });
		expect(commentCounts(deck(), 3)).toStrictEqual({ slide: 0, total: 3 });
	});
});

function makeHost(slides = deck()) {
	const log: string[] = [];
	const state = { slides, current: 0 };
	const host: ReviewCommentsHost = {
		slides: () => state.slides,
		currentSlide: () => state.current,
		goToSlide: (i) => {
			state.current = i;
			log.push(`go:${i}`);
		},
		ensureCommentsOpen: () => log.push('open'),
		applyComments: (changes: readonly SlideCommentsChange[]) => {
			log.push(`apply:${changes.map((x) => `${x.slideIndex}=${x.comments.length}`).join(',')}`);
			for (const change of changes) {
				state.slides = state.slides.map((s, i) =>
					i === change.slideIndex ? { ...s, comments: change.comments } : s,
				);
			}
		},
		root: () => document,
	};
	return { host, log, state };
}

describe('runReviewCommentCommand', () => {
	afterEach(() => {
		document.body.innerHTML = '';
	});

	it('moves to the next comment on another slide, opening the pane first', () => {
		document.body.innerHTML = '<ul><li data-comment-id="d" tabindex="0">d</li></ul>';
		const { host, log } = makeHost();
		expect(runReviewCommentCommand('review.comments.next', host, 'b')).toBe('d');
		expect(log).toStrictEqual(['open', 'go:2']);
		expect(
			document.querySelector('[data-comment-id="d"]')?.getAttribute('data-comment-active'),
		).toBe('true');
	});

	it('does nothing without comments', () => {
		const { host, log } = makeHost([{}, {}] as ReturnType<typeof deck>);
		expect(runReviewCommentCommand('review.comments.next', host, null)).toBeUndefined();
		expect(runReviewCommentCommand('review.comments.deleteAll', host, null)).toBeUndefined();
		expect(log).toStrictEqual([]);
	});

	it('deletes the current comment in one write and clears the cursor', () => {
		const { host, log, state } = makeHost();
		expect(runReviewCommentCommand('review.comments.delete', host, 'd')).toBeNull();
		expect(log).toStrictEqual(['apply:2=0']);
		expect(state.slides[2]?.comments).toStrictEqual([]);
		expect(runReviewCommentCommand('review.comments.delete', host, null)).toBeUndefined();
	});

	it('deletes the slide comments, keeping a cursor on another slide', () => {
		const { host, log } = makeHost();
		expect(runReviewCommentCommand('review.comments.deleteSlide', host, 'd')).toBe('d');
		expect(log).toStrictEqual(['apply:0=0']);
	});

	it('deletes every comment as a single write', () => {
		const { host, log } = makeHost();
		expect(runReviewCommentCommand('review.comments.deleteAll', host, 'a')).toBeNull();
		expect(log).toStrictEqual(['apply:0=0,2=0']);
	});
});

describe('comment rows', () => {
	afterEach(() => {
		document.body.innerHTML = '';
	});

	it('marks one row active and reports the row a click lands in', () => {
		document.body.innerHTML =
			'<div data-comment-id="a"><button id="x">x</button></div><div data-comment-id="b"></div>';
		expect(revealComment(document, 'b')).toBe(true);
		expect(document.querySelectorAll('[data-comment-active]')).toHaveLength(1);
		expect(revealComment(document, 'missing')).toBe(false);
		expect(document.querySelectorAll('[data-comment-active]')).toHaveLength(0);
		const seen: string[] = [];
		const stop = trackActiveComment(document, (id) => seen.push(id));
		document.getElementById('x')?.click();
		stop();
		document.getElementById('x')?.click();
		expect(seen).toStrictEqual(['a']);
	});
});
