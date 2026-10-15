/**
 * Review > Comments commands (Previous, Next, Delete, Delete Slide Comments, Delete All Comments)
 * as one framework-free runner every binding calls from its `command-request` handler.
 *
 * The binding supplies a {@link ReviewCommentsHost} (its slides, its current slide, a slide
 * navigator, a way to open its comments pane and one history-aware writer) and keeps the id of
 * the current comment; the runner decides what to do and returns the new current comment.
 *
 * @module render/review-comment-commands
 */
import type { PptxComment } from 'ooxml-core/pptx';

import type { CommentLocation, SlideWithComments } from './comment-navigation';
import {
	adjacentComment,
	clearAllComments,
	clearSlideComments,
	findCommentLocation,
	removeCommentFromSlides,
	revealComment,
} from './comment-navigation';

/** One slide whose comment list changed. */
export interface SlideCommentsChange {
	slideIndex: number;
	comments: PptxComment[];
}

export interface ReviewCommentsHost {
	slides(): readonly SlideWithComments[];
	currentSlide(): number;
	goToSlide(slideIndex: number): void;
	/** Open the comments pane if it is closed (never toggles it shut). */
	ensureCommentsOpen(): void;
	/** Write the changed slides as ONE history entry. */
	applyComments(changes: readonly SlideCommentsChange[]): void;
	/** Where the comment rows render; defaults to the host document. */
	root?(): ParentNode | null;
}

export type ReviewCommentCommandId =
	| 'review.comments.previous'
	| 'review.comments.next'
	| 'review.comments.delete'
	| 'review.comments.deleteSlide'
	| 'review.comments.deleteAll';

export const REVIEW_COMMENT_COMMAND_IDS: ReadonlySet<string> = new Set<ReviewCommentCommandId>([
	'review.comments.previous',
	'review.comments.next',
	'review.comments.delete',
	'review.comments.deleteSlide',
	'review.comments.deleteAll',
]);

export function isReviewCommentCommand(id: string): id is ReviewCommentCommandId {
	return REVIEW_COMMENT_COMMAND_IDS.has(id);
}

/** Comments on the current slide and in the whole deck, for the ribbon state. */
export function commentCounts(
	slides: readonly SlideWithComments[],
	currentSlide: number,
): { slide: number; total: number } {
	return {
		slide: slides[currentSlide]?.comments?.length ?? 0,
		total: slides.reduce((sum, slide) => sum + (slide.comments?.length ?? 0), 0),
	};
}

/** Retry the reveal for a few frames: the pane and the slide's rows render after the state change. */
function revealSoon(host: ReviewCommentsHost, id: string): void {
	const view = (host.root?.() ??
		(typeof document === 'undefined' ? null : document)) as ParentNode | null;
	if (!view) {
		return;
	}
	let attempts = 0;
	const attempt = (): void => {
		attempts += 1;
		if (!revealComment(view, id) && attempts < 60 && typeof requestAnimationFrame === 'function') {
			requestAnimationFrame(attempt);
		}
	};
	attempt();
}

function changesFor(
	before: readonly SlideWithComments[],
	after: readonly SlideWithComments[],
): SlideCommentsChange[] {
	const changes: SlideCommentsChange[] = [];
	after.forEach((slide, slideIndex) => {
		if (slide.comments !== before[slideIndex]?.comments) {
			changes.push({ slideIndex, comments: slide.comments ?? [] });
		}
	});
	return changes;
}

/**
 * Run `id`. Returns the id of the comment that is current afterwards (null when there is none),
 * or `undefined` when the command did nothing.
 */
export function runReviewCommentCommand(
	id: ReviewCommentCommandId,
	host: ReviewCommentsHost,
	activeId: string | null,
): string | null | undefined {
	const slides = host.slides();
	switch (id) {
		case 'review.comments.previous':
		case 'review.comments.next': {
			const target: CommentLocation | null = adjacentComment(
				slides,
				findCommentLocation(slides, activeId),
				host.currentSlide(),
				id === 'review.comments.next' ? 1 : -1,
			);
			if (!target) {
				return undefined;
			}
			host.ensureCommentsOpen();
			if (target.slideIndex !== host.currentSlide()) {
				host.goToSlide(target.slideIndex);
			}
			revealSoon(host, target.commentId);
			return target.commentId;
		}
		case 'review.comments.delete': {
			const next = activeId ? removeCommentFromSlides(slides, activeId) : null;
			if (!next) {
				return undefined;
			}
			host.applyComments(changesFor(slides, next));
			return null;
		}
		case 'review.comments.deleteSlide': {
			const next = clearSlideComments(slides, host.currentSlide());
			if (!next) {
				return undefined;
			}
			host.applyComments(changesFor(slides, next));
			return findCommentLocation(next, activeId) ? activeId : null;
		}
		case 'review.comments.deleteAll': {
			const next = clearAllComments(slides);
			if (!next) {
				return undefined;
			}
			host.applyComments(changesFor(slides, next));
			return null;
		}
	}
}
