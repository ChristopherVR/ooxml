import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { PptxCommentMention } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef, useState } from 'react';

import type { UseCommentsInput, UseCommentsResult } from './useComments-helpers';
import { useController } from './useController';

/**
 * The comments panel's state for React: the shared `createCommentsPanelController` read through
 * `useController`, mapped onto the handler names the inspector's comment components take.
 * Drafts, edits and reply composers live in the controller, per slide; the commands commit
 * through the viewer's `onUpdateSlides` and mark the document dirty.
 */
export function useComments(input: UseCommentsInput): UseCommentsResult {
	const latest = useRef(input);
	latest.current = input;
	const [controller] = useState(() =>
		createCommentsPanelController({
			slides: () => latest.current.slides,
			canEdit: () => latest.current.canEdit,
			authorName: () => latest.current.userName ?? 'You',
			selectedElementId: () => latest.current.selectedElementId,
			commit: (slides) => {
				latest.current.onUpdateSlides(() => slides);
				latest.current.onMarkDirty();
			},
		}),
	);
	const state = useController(controller);

	// Drafts for slides that no longer exist go with them. `input.slides` is the trigger; the
	// controller reads the deck itself.
	useEffect(() => {
		controller.pruneDrafts();
		// oxlint-disable-next-line react/exhaustive-effect-dependencies -- see comment above
	}, [controller, input.slides]);

	const handlers = useMemo(
		() => ({
			handleCommentDraftChange: controller.setDraft,
			handleAddSlideComment: (slideIndex: number) => {
				controller.add(slideIndex);
			},
			handleDeleteSlideComment: (slideIndex: number, commentId: string) => {
				controller.remove(slideIndex, commentId);
			},
			handleStartCommentEdit: controller.startEdit,
			handleCancelCommentEdit: controller.cancelEdit,
			handleSaveCommentEdit: (slideIndex: number, commentId: string) => {
				controller.saveEdit(slideIndex, commentId);
			},
			handleSetCommentEditDraft: controller.setEditDraft,
			handleToggleCommentResolved: (slideIndex: number, commentId: string) => {
				controller.toggleResolved(slideIndex, commentId);
			},
			handleStartReply: (_slideIndex: number, commentId: string) =>
				controller.startReply(commentId),
			handleCancelReply: controller.cancelReply,
			handleReplyDraftChange: controller.setReplyDraft,
			handleSubmitReply: (slideIndex: number, commentId: string) => {
				controller.submitReply(slideIndex, commentId);
			},
		}),
		[controller],
	);

	return {
		commentDraftBySlideId: state.draftBySlideId,
		editingCommentIdBySlideId: state.editingIdBySlideId,
		commentEditDraftByCommentId: state.editDraftById,
		// The controller never mutates these arrays; the inspector's props are typed mutable.
		commentDraftMentionsBySlideId: state.draftMentionsBySlideId as Record<
			string,
			PptxCommentMention[]
		>,
		replyDraftMentionsByCommentId: state.replyMentionsById as Record<string, PptxCommentMention[]>,
		replyingToCommentId: state.replyingToId,
		replyDraftByCommentId: state.replyDraftById,
		...handlers,
	};
}
