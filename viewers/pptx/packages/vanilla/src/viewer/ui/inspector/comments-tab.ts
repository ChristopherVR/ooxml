import type { PptxCommentMention } from 'pptx-viewer-core';

import type { CommentsPanelBinding } from '../../editor/editor-comment-actions';
import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { attachCommentMentionTypeahead } from '../comment-mention-typeahead';
import { createCommentThreadView } from '../comment-thread-view';
import type { InspectorDeckState, InspectorHandlers } from './types';

export interface CommentsTab {
	el: HTMLElement;
	update(state: InspectorDeckState): void;
}

type CommentsHandlers = Pick<InspectorHandlers, 'commentsPanel'>;

/**
 * The inspector's Comments tab: current-slide comment threads with nested
 * replies, edit-in-place, resolve and delete, plus an add box in editable
 * mode (port of React's `InspectorCommentsSection` / `InspectorCommentRow`).
 *
 * The thread list itself is `createCommentThreadView`, shared with the
 * workspace Comments pane so vanilla's two comment surfaces cannot offer
 * different affordances, and both read the viewer's shared comments panel
 * controller, so the draft typed here is the one the pane shows.
 */
export function createCommentsTab(
	doc: Document,
	t: Translator,
	handlers: CommentsHandlers,
): CommentsTab {
	const el = createEl(doc, 'div', 'pptxv-inspector-comments');
	let mentionAuthors: InspectorDeckState['commentMentionAuthors'] = [];
	// Read lazily: the editor's actions exist only once the viewer has wired them.
	const panel: CommentsPanelBinding = {
		get controller() {
			return handlers.commentsPanel().controller;
		},
		slideIndex: () => handlers.commentsPanel().slideIndex(),
		slideId: () => handlers.commentsPanel().slideId(),
		canEdit: () => handlers.commentsPanel().canEdit(),
	};
	const threads = createCommentThreadView(doc, t, panel, () => mentionAuthors);
	el.appendChild(threads.el);
	const empty = createEl(doc, 'p', 'pptxv-inspector-empty');
	empty.textContent = t('pptx.comments.noComments');
	el.appendChild(empty);

	const addBox = createEl(doc, 'div', 'pptxv-inspector-comment-add');
	const input = doc.createElement('textarea');
	input.className = 'pptxv-inspector-comment-input';
	input.rows = 2;
	input.placeholder = t('pptx.comments.addCommentPlaceholder');
	input.setAttribute('aria-label', t('pptx.comments.addComment'));
	const draftMentions = (): PptxCommentMention[] => [
		...(panel.controller.getSnapshot().draftMentionsBySlideId[panel.slideId()] ?? []),
	];
	input.addEventListener('input', () => panel.controller.setDraft(panel.slideId(), input.value));
	const draftMentionTypeahead = attachCommentMentionTypeahead({
		doc,
		t,
		field: input,
		getAuthors: () => mentionAuthors,
		getMentions: draftMentions,
		onChange: (next) => {
			input.value = next.text;
			panel.controller.setDraft(panel.slideId(), next.text, next.mentions);
		},
	});
	const addBtn = createEl(doc, 'button', 'pptxv-inspector-deck-btn');
	addBtn.type = 'button';
	addBtn.textContent = t('pptx.comments.addComment');
	addBtn.addEventListener('click', () => {
		if (panel.controller.add(panel.slideIndex())) {
			input.value = '';
		}
	});
	addBox.append(input, draftMentionTypeahead.el, addBtn);
	el.appendChild(addBox);

	return {
		el,
		update(state) {
			mentionAuthors = state.commentMentionAuthors;
			// The slide's own draft (a slide change brings its draft with it); not while typing.
			if (input.ownerDocument.activeElement !== input) {
				input.value = panel.controller.getSnapshot().draftBySlideId[panel.slideId()] ?? '';
			}
			empty.hidden = state.comments.length > 0;
			addBox.hidden = !state.editable;
			threads.update(state.comments, state.editable);
		},
	};
}
