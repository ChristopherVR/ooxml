/**
 * Pure, framework-agnostic `@`-mention logic for comment bodies.
 *
 * A comment's text is a flat string; its mentions are offset spans into that
 * string (see `PptxCommentMention`). Every binding needs the same two answers:
 *
 * - DISPLAY: "which slices of this string are mentions?" -> `commentTextSegments`
 *   returns a framework-neutral `CommentTextSegment[]` that a template maps
 *   straight onto spans.
 * - AUTHORING: "the user picked an author from the `@` typeahead"
 *   -> `insertCommentMention` returns the new text, the new mention list, and
 *   where the caret should land.
 *
 * The logic is the format-neutral one in `comments/mentions` (which the shared comments pane uses);
 * this module maps PowerPoint's mention records onto it. No framework imports.
 */

import type { PptxCommentMention, PptxModernCommentAuthor } from 'ooxml-core/pptx';

import {
	commentSegments,
	insertMention,
	matchMentionCandidates,
	mentionQueryAt,
	type OfficeCommentMention,
} from '../../comments/mentions';

/** DOM attribute every binding stamps on a rendered mention, for e2e. */
export const COMMENT_MENTION_ATTRIBUTE = 'data-pptx-comment-mention';

/** CSS class every binding puts on a rendered mention span. */
export const COMMENT_MENTION_CLASS = 'pptx-comment-mention';

/** One run of a comment body: either inert text or an `@`-mention. */
export interface CommentTextSegment {
	kind: 'text' | 'mention';
	text: string;
	/** Display name of the mentioned author, when it resolved. */
	authorName?: string;
	/** `p188:author` id of the mentioned person. */
	personId?: string;
}

const toSpan = (mention: PptxCommentMention): OfficeCommentMention => ({
	start: mention.startIndex,
	length: mention.length,
	...(mention.personId !== undefined ? { id: mention.personId } : {}),
	...(mention.authorName !== undefined ? { name: mention.authorName } : {}),
});

/**
 * Split a comment body into plain-text and mention segments.
 *
 * Overlapping, out-of-range and zero-length spans are ignored rather than
 * producing garbled output, because the offsets come from a foreign file and
 * an edit made outside this library can invalidate them.
 */
export function commentTextSegments(
	text: string,
	mentions?: PptxCommentMention[],
): CommentTextSegment[] {
	return commentSegments(text, mentions?.map(toSpan)).map((segment) => ({
		kind: segment.kind,
		text: segment.text,
		...(segment.name !== undefined ? { authorName: segment.name } : {}),
		...(segment.id !== undefined ? { personId: segment.id } : {}),
	}));
}

/** Whether this body has at least one renderable mention. */
export function hasCommentMentions(text: string, mentions?: PptxCommentMention[]): boolean {
	return commentTextSegments(text, mentions).some((segment) => segment.kind === 'mention');
}

/** An in-progress `@` token the author is typing, for the mention typeahead. */
export interface CommentMentionQuery {
	/** Text typed after the `@`, lower-cased for matching. */
	query: string;
	/** Index of the `@` itself. */
	start: number;
}

/**
 * The `@`-token immediately left of `caret`, or `null` when there is none.
 *
 * The `@` must start the body or follow whitespace, and the token must not
 * contain whitespace, so an email address does not open the picker.
 */
export function commentMentionQuery(text: string, caret: number): CommentMentionQuery | null {
	return mentionQueryAt(text, caret);
}

/** Authors whose name or initials match a typeahead query, best-first. */
export function matchCommentMentionAuthors(
	authors: PptxModernCommentAuthor[],
	query: string,
	limit = 6,
): PptxModernCommentAuthor[] {
	return matchMentionCandidates(authors, query, limit);
}

/** Result of accepting an author from the mention typeahead. */
export interface CommentMentionInsertion {
	text: string;
	mentions: PptxCommentMention[];
	/** Where the caret belongs after the insertion. */
	caret: number;
}

/** A brace-wrapped GUID, the shape `mentionId` uses. */
export function defaultMentionId(): string {
	if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
		return `{${crypto.randomUUID().toUpperCase()}}`;
	}
	const hex = (length: number): string =>
		Array.from({ length }, () =>
			Math.floor(Math.random() * 16)
				.toString(16)
				.toUpperCase(),
		).join('');
	return `{${hex(8)}-${hex(4)}-4${hex(3)}-8${hex(3)}-${hex(12)}}`;
}

/**
 * Replace the `@`-token at `caret` with `@<author name>` and record the mention.
 *
 * Existing mentions after the insertion point are shifted by the length delta,
 * so a body with several mentions stays consistent while it is being typed.
 * The inserted span INCLUDES the `@`, which is what Office records.
 */
export function insertCommentMention(
	text: string,
	mentions: PptxCommentMention[] | undefined,
	caret: number,
	author: PptxModernCommentAuthor,
	generateId: () => string = defaultMentionId,
): CommentMentionInsertion {
	const existing = mentions ?? [];
	// Each existing record is tagged with its index, so the shifted span finds its record again.
	const tagged = existing.map((mention, index) => ({
		start: mention.startIndex,
		length: mention.length,
		id: String(index),
	}));
	const result = insertMention(text, tagged, caret, { id: author.id, name: author.name });
	return {
		text: result.text,
		caret: result.caret,
		mentions: result.mentions.map((span) =>
			span.name === undefined
				? { ...existing[Number(span.id)]!, startIndex: span.start }
				: {
						id: generateId(),
						personId: author.id,
						authorName: author.name,
						startIndex: span.start,
						length: span.length,
					},
		),
	};
}
