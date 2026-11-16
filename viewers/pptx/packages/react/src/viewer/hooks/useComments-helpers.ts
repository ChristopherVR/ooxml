// The input and result shapes of `useComments`. Its state and commands are the shared
// `createCommentsPanelController` (`ooxml-ui/pptx`); the comment-array transforms are the shared
// `comments-list`.
import type { PptxCommentMention, PptxSlide } from 'pptx-viewer-core';

export interface UseCommentsInput {
	slides: PptxSlide[];
	activeSlideIndex: number;
	canEdit: boolean;
	userName?: string;
	selectedElementId?: string | null;
	onUpdateSlides: (updater: (slides: PptxSlide[]) => PptxSlide[]) => void;
	onMarkDirty: () => void;
}

export interface UseCommentsResult {
	commentDraftBySlideId: Record<string, string>;
	editingCommentIdBySlideId: Record<string, string | null>;
	commentEditDraftByCommentId: Record<string, string>;
	replyingToCommentId: string | null;
	replyDraftByCommentId: Record<string, string>;
	/** `@`-mentions accumulated on the new-comment draft for the active slide. */
	commentDraftMentionsBySlideId: Record<string, PptxCommentMention[]>;
	/** `@`-mentions accumulated on each reply draft, keyed by parent comment id. */
	replyDraftMentionsByCommentId: Record<string, PptxCommentMention[]>;
	/**
	 * `mentions` is passed through unchanged on a plain text edit (only
	 * accepting a typeahead suggestion updates it); omitting it keeps
	 * whatever was accumulated so far.
	 */
	handleCommentDraftChange: (
		slideId: string,
		draft: string,
		mentions?: PptxCommentMention[],
	) => void;
	handleAddSlideComment: (slideIndex: number) => void;
	handleDeleteSlideComment: (slideIndex: number, commentId: string) => void;
	handleStartCommentEdit: (slideId: string, commentId: string) => void;
	handleCancelCommentEdit: (slideId: string) => void;
	handleSaveCommentEdit: (slideIndex: number, commentId: string) => void;
	handleSetCommentEditDraft: (commentId: string, draft: string) => void;
	handleToggleCommentResolved: (slideIndex: number, commentId: string) => void;
	handleStartReply: (slideIndex: number, commentId: string) => void;
	handleCancelReply: () => void;
	handleReplyDraftChange: (
		commentId: string,
		draft: string,
		mentions?: PptxCommentMention[],
	) => void;
	handleSubmitReply: (slideIndex: number, commentId: string) => void;
}
