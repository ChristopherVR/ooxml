/**
 * Review > Comments navigation and bulk delete: Previous / Next across slides in document
 * order, Delete (the current comment), Delete All Comments on This Slide / in This Presentation.
 *
 * Pure functions over the deck's slides, shared by every binding; a binding only supplies its
 * slides, its current slide, a slide navigator and its history-aware slide writer. The
 * "current comment" cursor is DOM-derived (`trackActiveComment`) so each binding's own comments
 * panel needs only to tag its rows with `data-comment-id`.
 *
 * @module render/comment-navigation
 */
import type { PptxComment } from 'ooxml-core/pptx';

import { removeCommentFromList } from './comments-list';

/** The slide fields comment navigation reads and rewrites. */
export interface SlideWithComments {
	comments?: PptxComment[] | undefined;
}

export interface CommentLocation {
	slideIndex: number;
	commentId: string;
}

/** Top-level comments (a thread is one stop) in slide order, then comment order. */
export function listCommentLocations(slides: readonly SlideWithComments[]): CommentLocation[] {
	const out: CommentLocation[] = [];
	slides.forEach((slide, slideIndex) => {
		for (const comment of slide.comments ?? []) {
			out.push({ slideIndex, commentId: comment.id });
		}
	});
	return out;
}

/** The location of comment `id` (a reply resolves to its thread's top-level comment). */
export function findCommentLocation(
	slides: readonly SlideWithComments[],
	id: string | null | undefined,
): CommentLocation | null {
	if (!id) {
		return null;
	}
	for (let slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
		for (const comment of slides[slideIndex]?.comments ?? []) {
			if (comment.id === id || containsReply(comment, id)) {
				return { slideIndex, commentId: comment.id };
			}
		}
	}
	return null;
}

function containsReply(comment: PptxComment, id: string): boolean {
	return (comment.replies ?? []).some((reply) => reply.id === id || containsReply(reply, id));
}

/**
 * The comment after (`1`) or before (`-1`) `current`, wrapping around the deck like PowerPoint.
 * With no current comment, Next goes to the first one on or after `currentSlide` and Previous to
 * the last one on or before it. Null when the deck has no comments.
 */
export function adjacentComment(
	slides: readonly SlideWithComments[],
	current: CommentLocation | null,
	currentSlide: number,
	direction: 1 | -1,
): CommentLocation | null {
	const all = listCommentLocations(slides);
	if (all.length === 0) {
		return null;
	}
	const at = current
		? all.findIndex(
				(item) => item.slideIndex === current.slideIndex && item.commentId === current.commentId,
			)
		: -1;
	if (at >= 0) {
		return all[(at + direction + all.length) % all.length] ?? null;
	}
	if (direction === 1) {
		return all.find((item) => item.slideIndex >= currentSlide) ?? all[0] ?? null;
	}
	return [...all].reverse().find((item) => item.slideIndex <= currentSlide) ?? all.at(-1) ?? null;
}

/** `slides` without comment `id` (any depth), or null when it is not there. */
export function removeCommentFromSlides<T extends SlideWithComments>(
	slides: readonly T[],
	id: string,
): T[] | null {
	for (let slideIndex = 0; slideIndex < slides.length; slideIndex += 1) {
		const slide = slides[slideIndex];
		const next = slide?.comments ? removeCommentFromList(slide.comments, id) : null;
		if (slide && next) {
			return slides.map((candidate, index) =>
				index === slideIndex ? { ...candidate, comments: next } : candidate,
			);
		}
	}
	return null;
}

/** `slides` with every comment on slide `slideIndex` removed, or null when it had none. */
export function clearSlideComments<T extends SlideWithComments>(
	slides: readonly T[],
	slideIndex: number,
): T[] | null {
	if (!slides[slideIndex]?.comments?.length) {
		return null;
	}
	return slides.map((slide, index) => (index === slideIndex ? { ...slide, comments: [] } : slide));
}

/** `slides` with every comment in the presentation removed, or null when there were none. */
export function clearAllComments<T extends SlideWithComments>(slides: readonly T[]): T[] | null {
	if (!slides.some((slide) => slide.comments?.length)) {
		return null;
	}
	return slides.map((slide) => (slide.comments?.length ? { ...slide, comments: [] } : slide));
}

/** Attribute every binding's comments panel tags a comment row with. */
export const COMMENT_ID_ATTR = 'data-comment-id';
/** Set on the row of the comment the ribbon cursor is on. */
export const COMMENT_ACTIVE_ATTR = 'data-comment-active';

const ACTIVE_STYLE_ID = 'pptx-comment-active-style';

/** One shared rule so every binding's row shows the current comment the same way. */
function ensureActiveStyle(doc: Document | null): void {
	if (!doc || doc.getElementById(ACTIVE_STYLE_ID) || !doc.head) {
		return;
	}
	const style = doc.createElement('style');
	style.id = ACTIVE_STYLE_ID;
	style.textContent = `[${COMMENT_ACTIVE_ATTR}]{outline:2px solid #2563eb;outline-offset:1px;}`;
	doc.head.append(style);
}

function escapeAttr(value: string): string {
	return value.replace(/["\\]/gu, '\\$&');
}

/**
 * Scroll the row of comment `id` into view, focus it and mark it active (clearing the mark on
 * every other row). Returns false when no row is rendered (the panel is closed).
 */
export function revealComment(root: ParentNode, id: string): boolean {
	for (const stale of root.querySelectorAll(`[${COMMENT_ACTIVE_ATTR}]`)) {
		stale.removeAttribute(COMMENT_ACTIVE_ATTR);
	}
	const row = root.querySelector<HTMLElement>(`[${COMMENT_ID_ATTR}="${escapeAttr(id)}"]`);
	if (!row) {
		return false;
	}
	ensureActiveStyle(row.ownerDocument);
	row.setAttribute(COMMENT_ACTIVE_ATTR, 'true');
	if (!row.hasAttribute('tabindex')) {
		row.setAttribute('tabindex', '-1');
	}
	row.scrollIntoView?.({ block: 'nearest' });
	row.focus({ preventScroll: true });
	return true;
}

/**
 * Report the comment a click or focus lands on (any element inside a tagged row), so the ribbon
 * cursor follows the user's own picks in the panel. Returns the disposer.
 */
export function trackActiveComment(
	root: Document | ShadowRoot,
	onChange: (commentId: string) => void,
): () => void {
	const handler = (event: Event): void => {
		const target = event.target;
		const row =
			target instanceof Element ? target.closest<HTMLElement>(`[${COMMENT_ID_ATTR}]`) : null;
		const id = row?.getAttribute(COMMENT_ID_ATTR);
		if (id) {
			onChange(id);
		}
	};
	root.addEventListener('focusin', handler);
	root.addEventListener('click', handler);
	return () => {
		root.removeEventListener('focusin', handler);
		root.removeEventListener('click', handler);
	};
}
