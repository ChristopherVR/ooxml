/**
 * The comments panel as one state owner for every binding: the new-comment draft and its
 * `@`-mentions per slide, the comment being edited per slide and its edit draft, the open reply
 * composer and each reply draft with its mentions, and the add, edit, delete, resolve and reply
 * commands that write a slide's comments back through the binding's undoable commit.
 *
 * The comment-array transforms were already shared (`comments-list`); the panels around them were
 * not, and four of five kept less state than React's: Vue, Angular and Svelte had one draft for the
 * whole deck (switching slides carried a half-typed comment onto the next slide) and no edit in
 * place, and only React anchored a new comment to the selected element. React's behaviour is the one
 * chosen: drafts and edits are kept per slide, a reply composer opens empty, and a draft for a
 * slide that was deleted is dropped.
 *
 * A binding keeps one controller per viewer (drafts outlive the panel), reads `getSnapshot()`
 * through its store adapter and keeps only its templates and its mention typeahead, which reports
 * the text and mentions to {@link CommentsPanelController.setDraft} and friends.
 *
 * @module render/controller/comments-panel-controller
 */
import type { PptxComment, PptxCommentMention, PptxSlide } from 'ooxml-core/pptx';

import {
	addCommentToList,
	editCommentInList,
	removeCommentFromList,
	replyToCommentInList,
	toggleCommentResolvedInList,
} from '../comments-list';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

export interface CommentsPanelState {
	/** The new-comment draft of each slide, by slide id, exactly as typed. */
	readonly draftBySlideId: Readonly<Record<string, string>>;
	/** The `@`-mentions recorded in each slide's draft. */
	readonly draftMentionsBySlideId: Readonly<Record<string, readonly PptxCommentMention[]>>;
	/** The comment being edited on each slide, or null. */
	readonly editingIdBySlideId: Readonly<Record<string, string | null>>;
	/** The edit draft of each comment being (or last) edited, by comment id. */
	readonly editDraftById: Readonly<Record<string, string>>;
	/** The comment whose reply composer is open (one at a time), or null. */
	readonly replyingToId: string | null;
	/** Each reply draft, by parent comment id. */
	readonly replyDraftById: Readonly<Record<string, string>>;
	/** The `@`-mentions recorded in each reply draft. */
	readonly replyMentionsById: Readonly<Record<string, readonly PptxCommentMention[]>>;
}

export interface CommentsPanelControllerDeps {
	/** The deck as it is now. */
	slides(): readonly PptxSlide[];
	/** Whether the host allows edits; every command is refused otherwise. */
	canEdit(): boolean;
	/** The author written on a new comment or reply. */
	authorName(): string;
	/** The element a new comment is anchored to (the selection), if any. */
	selectedElementId?(): string | null | undefined;
	/** Replace the deck: one undoable step, marked dirty. */
	commit(slides: PptxSlide[]): void;
}

export interface CommentsPanelController extends PptxController<CommentsPanelState> {
	/** The new-comment draft changed. Omitted `mentions` keep the ones recorded so far. */
	setDraft(slideId: string, text: string, mentions?: readonly PptxCommentMention[]): void;
	/** Add the slide's draft as a comment; returns the new comment's id, or null. */
	add(slideIndex: number): string | null;
	remove(slideIndex: number, commentId: string): boolean;
	toggleResolved(slideIndex: number, commentId: string): boolean;
	/** Edit a comment (or a reply) in place, starting from its text. */
	startEdit(slideId: string, commentId: string): void;
	cancelEdit(slideId: string): void;
	setEditDraft(commentId: string, text: string): void;
	/** Save the edit draft (trimmed; a blank one is refused). */
	saveEdit(slideIndex: number, commentId: string): boolean;
	/** Open the reply composer under `commentId`, empty. */
	startReply(commentId: string): void;
	cancelReply(): void;
	setReplyDraft(commentId: string, text: string, mentions?: readonly PptxCommentMention[]): void;
	/** Post the reply draft under `commentId`; returns the reply's id, or null. */
	submitReply(slideIndex: number, commentId: string): string | null;
	/** Drop the drafts and edits of slides no longer in the deck (call when the deck changes). */
	pruneDrafts(): void;
}

export const INITIAL_COMMENTS_PANEL_STATE: CommentsPanelState = Object.freeze({
	draftBySlideId: Object.freeze({}),
	draftMentionsBySlideId: Object.freeze({}),
	editingIdBySlideId: Object.freeze({}),
	editDraftById: Object.freeze({}),
	replyingToId: null,
	replyDraftById: Object.freeze({}),
	replyMentionsById: Object.freeze({}),
});

/** Find a comment or a reply at any depth. */
export function findCommentInThread(
	comments: readonly PptxComment[],
	id: string,
): PptxComment | undefined {
	for (const comment of comments) {
		if (comment.id === id) {
			return comment;
		}
		const nested = findCommentInThread(comment.replies ?? [], id);
		if (nested) {
			return nested;
		}
	}
	return undefined;
}

/** `slides` with slide `slideIndex`'s comments replaced, every other slide untouched. */
export function replaceSlideComments(
	slides: readonly PptxSlide[],
	slideIndex: number,
	comments: PptxComment[],
): PptxSlide[] {
	return slides.map((slide, index) => (index === slideIndex ? { ...slide, comments } : slide));
}

function without<T>(map: Readonly<Record<string, T>>, key: string): Readonly<Record<string, T>> {
	if (!(key in map)) {
		return map;
	}
	const next = { ...map };
	delete next[key];
	return next;
}

function keepKeys<T>(
	map: Readonly<Record<string, T>>,
	keys: ReadonlySet<string>,
): Readonly<Record<string, T>> {
	const entries = Object.entries(map);
	const kept = entries.filter(([key]) => keys.has(key));
	return kept.length === entries.length ? map : Object.fromEntries(kept);
}

export function createCommentsPanelController(
	deps: CommentsPanelControllerDeps,
): CommentsPanelController {
	const store = createControllerStore<CommentsPanelState>(INITIAL_COMMENTS_PANEL_STATE);
	const patch = (next: Partial<CommentsPanelState>): void =>
		store.update((current) => ({ ...current, ...next }));

	/** Run `transform` over slide `slideIndex`'s comments and commit a change. */
	const apply = (
		slideIndex: number,
		transform: (comments: PptxComment[]) => PptxComment[] | null,
	): PptxComment[] | null => {
		if (!deps.canEdit()) {
			return null;
		}
		const slides = deps.slides();
		const slide = slides[slideIndex];
		if (!slide) {
			return null;
		}
		const next = transform([...(slide.comments ?? [])]);
		if (!next) {
			return null;
		}
		deps.commit(replaceSlideComments(slides, slideIndex, next));
		return next;
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		setDraft(slideId, text, mentions) {
			const state = store.getSnapshot();
			patch({
				draftBySlideId: { ...state.draftBySlideId, [slideId]: text },
				...(mentions
					? { draftMentionsBySlideId: { ...state.draftMentionsBySlideId, [slideId]: mentions } }
					: {}),
			});
		},
		add(slideIndex) {
			const slide = deps.slides()[slideIndex];
			if (!slide) {
				return null;
			}
			const state = store.getSnapshot();
			// Untrimmed: the mentions index the draft as typed, and the shared add trims both.
			const text = state.draftBySlideId[slide.id] ?? '';
			if (text.trim().length === 0) {
				return null;
			}
			const next = apply(slideIndex, (comments) =>
				addCommentToList(
					comments,
					text,
					deps.authorName(),
					undefined,
					undefined,
					deps.selectedElementId?.() ?? undefined,
					state.draftMentionsBySlideId[slide.id],
				),
			);
			if (!next) {
				return null;
			}
			const latest = store.getSnapshot();
			patch({
				draftBySlideId: { ...latest.draftBySlideId, [slide.id]: '' },
				draftMentionsBySlideId: without(latest.draftMentionsBySlideId, slide.id),
			});
			return next.at(-1)?.id ?? null;
		},
		remove(slideIndex, commentId) {
			const slide = deps.slides()[slideIndex];
			const next = apply(slideIndex, (comments) => removeCommentFromList(comments, commentId));
			if (!next || !slide) {
				return false;
			}
			const latest = store.getSnapshot();
			patch({ editingIdBySlideId: { ...latest.editingIdBySlideId, [slide.id]: null } });
			return true;
		},
		toggleResolved(slideIndex, commentId) {
			return (
				apply(slideIndex, (comments) => toggleCommentResolvedInList(comments, commentId)) !== null
			);
		},
		startEdit(slideId, commentId) {
			const slide = deps.slides().find((candidate) => candidate.id === slideId);
			const comment = slide ? findCommentInThread(slide.comments ?? [], commentId) : undefined;
			if (!comment) {
				return;
			}
			const state = store.getSnapshot();
			patch({
				editingIdBySlideId: { ...state.editingIdBySlideId, [slideId]: commentId },
				editDraftById: { ...state.editDraftById, [commentId]: comment.text ?? '' },
			});
		},
		cancelEdit(slideId) {
			const state = store.getSnapshot();
			if ((state.editingIdBySlideId[slideId] ?? null) !== null) {
				patch({ editingIdBySlideId: { ...state.editingIdBySlideId, [slideId]: null } });
			}
		},
		setEditDraft(commentId, text) {
			const state = store.getSnapshot();
			patch({ editDraftById: { ...state.editDraftById, [commentId]: text } });
		},
		saveEdit(slideIndex, commentId) {
			const slide = deps.slides()[slideIndex];
			const text = (store.getSnapshot().editDraftById[commentId] ?? '').trim();
			if (!slide || text.length === 0) {
				return false;
			}
			const next = apply(slideIndex, (comments) => editCommentInList(comments, commentId, text));
			if (!next) {
				return false;
			}
			const latest = store.getSnapshot();
			patch({ editingIdBySlideId: { ...latest.editingIdBySlideId, [slide.id]: null } });
			return true;
		},
		startReply(commentId) {
			const state = store.getSnapshot();
			patch({
				replyingToId: commentId,
				replyDraftById: { ...state.replyDraftById, [commentId]: '' },
				replyMentionsById: without(state.replyMentionsById, commentId),
			});
		},
		cancelReply() {
			if (store.getSnapshot().replyingToId !== null) {
				patch({ replyingToId: null });
			}
		},
		setReplyDraft(commentId, text, mentions) {
			const state = store.getSnapshot();
			patch({
				replyDraftById: { ...state.replyDraftById, [commentId]: text },
				...(mentions
					? { replyMentionsById: { ...state.replyMentionsById, [commentId]: mentions } }
					: {}),
			});
		},
		submitReply(slideIndex, commentId) {
			const state = store.getSnapshot();
			// Untrimmed: the mentions index the draft as typed, and the shared reply trims both.
			const text = state.replyDraftById[commentId] ?? '';
			if (text.trim().length === 0) {
				return null;
			}
			const next = apply(slideIndex, (comments) =>
				replyToCommentInList(
					comments,
					commentId,
					text,
					deps.authorName(),
					state.replyMentionsById[commentId],
				),
			);
			if (!next) {
				return null;
			}
			const latest = store.getSnapshot();
			patch({
				replyingToId: null,
				replyDraftById: without(latest.replyDraftById, commentId),
				replyMentionsById: without(latest.replyMentionsById, commentId),
			});
			const parent = findCommentInThread(next, commentId);
			return parent?.replies?.at(-1)?.id ?? null;
		},
		pruneDrafts() {
			const ids = new Set(deps.slides().map((slide) => slide.id));
			const state = store.getSnapshot();
			const draftBySlideId = keepKeys(state.draftBySlideId, ids);
			const draftMentionsBySlideId = keepKeys(state.draftMentionsBySlideId, ids);
			const editingIdBySlideId = keepKeys(state.editingIdBySlideId, ids);
			if (
				draftBySlideId !== state.draftBySlideId ||
				draftMentionsBySlideId !== state.draftMentionsBySlideId ||
				editingIdBySlideId !== state.editingIdBySlideId
			) {
				patch({ draftBySlideId, draftMentionsBySlideId, editingIdBySlideId });
			}
		},
	};
}
