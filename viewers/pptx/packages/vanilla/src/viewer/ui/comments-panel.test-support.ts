import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { PptxComment, PptxSlide } from 'pptx-viewer-core';

import type { CommentsPanelBinding } from '../editor/editor-comment-actions';

/**
 * A comments panel binding over a real shared controller and a two-slide deck whose first slide
 * holds `comments`, for tests of the comment surfaces.
 */
export function createTestCommentsPanel(comments: PptxComment[], editable = true) {
	let slides: PptxSlide[] = [
		{ id: 's1', rId: 's1', slideNumber: 1, elements: [], comments },
		{ id: 's2', rId: 's2', slideNumber: 2, elements: [] },
	];
	let current = 0;
	const controller = createCommentsPanelController({
		slides: () => slides,
		canEdit: () => editable,
		authorName: () => 'You',
		commit: (next) => {
			slides = next;
		},
	});
	const panel: CommentsPanelBinding = {
		controller,
		slideIndex: () => current,
		slideId: () => slides[current]?.id ?? '',
		canEdit: () => editable,
	};
	return {
		panel,
		slides: () => slides,
		comments: () => slides[current]?.comments ?? [],
		setCurrent: (index: number) => {
			current = index;
		},
	};
}
