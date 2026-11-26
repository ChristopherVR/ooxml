import { attachCommentsTaskPane, focusCommentsTaskPane } from 'ooxml-ui/pptx';
import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor } from 'pptx-viewer-core';

import type { CommentsPanelBinding } from '../editor/editor-comment-actions';
import type { Translator } from '../i18n';
import { createEl } from '../render';
import { attachCommentMentionTypeahead } from './comment-mention-typeahead';
import { createCommentThreadView } from './comment-thread-view';

/** Live view of the active slide's comments backing the workspace pane. */
export interface CommentsPanelModel {
	/** The active slide's comments at this moment. */
	getComments(): readonly PptxComment[];
	/** Authors offered by the `@`-mention typeahead. */
	getAuthors(): readonly PptxModernCommentAuthor[];
	/** Notify on any state change; returns an unsubscribe function. */
	subscribe(listener: () => void): () => void;
}

/** Detach hook stored on the pane so a replacement can end its subscription. */
const DISPOSE = Symbol('pptxv-comments-panel-dispose');

interface DisposablePane extends HTMLElement {
	[DISPOSE]?: () => void;
}

/**
 * Close the docked comments pane and release its store subscription.
 */
export function closeCommentsPanel(host: HTMLElement): boolean {
	const pane = host.querySelector<DisposablePane>('[data-pptx-comments-panel]');
	if (!pane) return false;
	pane[DISPOSE]?.();
	pane.remove();
	return true;
}

export function openCommentsPanel(
	doc: Document,
	host: HTMLElement,
	t: Translator,
	model: CommentsPanelModel,
	/** The viewer's shared comments panel state: drafts and commands. */
	panel: CommentsPanelBinding,
	onClose?: () => void,
): void {
	const existing = host.querySelector<DisposablePane>('[data-pptx-comments-panel]');
	if (existing) {
		focusCommentsTaskPane(existing);
		return;
	}

	const pane: DisposablePane = createEl(doc, 'aside', 'pptxv-workspace-pane');
	const header = createEl(doc, 'header');
	const title = createEl(doc, 'h2');
	title.textContent = t('pptx.toolbar.comments');
	header.append(title);
	pane.appendChild(header);

	const list = createEl(doc, 'div', 'pptxv-workspace-list');
	const threads = createCommentThreadView(doc, t, panel, model.getAuthors);
	const empty = createEl(doc, 'p');
	empty.textContent = t('pptx.comments.noneOnSlide');
	list.append(threads.el, empty);

	const render = (): void => {
		const comments: readonly PptxComment[] = model.getComments();
		empty.hidden = comments.length > 0;
		threads.update(comments, true);
	};

	// The new-comment draft is the slide's, kept by the shared controller.
	const draft = createEl(doc, 'textarea');
	draft.dataset.pptxCommentDraft = '';
	draft.placeholder = t('pptx.comments.addPlaceholder');
	const draftMentions = (): PptxCommentMention[] => [
		...(panel.controller.getSnapshot().draftMentionsBySlideId[panel.slideId()] ?? []),
	];
	let draftSlideId = panel.slideId();
	const showDraft = (): void => {
		draftSlideId = panel.slideId();
		draft.value = panel.controller.getSnapshot().draftBySlideId[draftSlideId] ?? '';
	};
	draft.addEventListener('input', () => panel.controller.setDraft(panel.slideId(), draft.value));
	const draftMentionTypeahead = attachCommentMentionTypeahead({
		doc,
		t,
		field: draft,
		getAuthors: model.getAuthors,
		getMentions: draftMentions,
		onChange: (next) => {
			draft.value = next.text;
			panel.controller.setDraft(panel.slideId(), next.text, next.mentions);
		},
	});
	const add = createEl(doc, 'button');
	add.type = 'button';
	add.textContent = t('pptx.comments.addComment');
	add.addEventListener('click', () => {
		if (panel.controller.add(panel.slideIndex())) {
			// The subscription re-renders the card list; the pane stays open.
			showDraft();
		}
	});
	showDraft();
	list.append(draft, draftMentionTypeahead.el, add);
	pane.appendChild(list);

	// Re-render only when the comment ARRAY changes (every mutation replaces
	// it), so unrelated store traffic never clobbers an in-progress card edit.
	let lastRendered = model.getComments();
	const dispose = (): void => {
		unsubscribe();
		draftMentionTypeahead.destroy();
		detach();
		delete pane[DISPOSE];
	};
	const unsubscribe = model.subscribe(() => {
		if (!pane.isConnected) {
			dispose();
			return;
		}
		if (panel.slideId() !== draftSlideId) {
			showDraft();
		}
		const next = model.getComments();
		if (next !== lastRendered) {
			lastRendered = next;
			render();
		}
	});
	pane[DISPOSE] = dispose;

	render();
	host.appendChild(pane);
	const detach = attachCommentsTaskPane(pane, {
		label: t('pptx.toolbar.comments'),
		closeLabel: t('pptx.common.close'),
		close: () => {
			dispose();
			pane.remove();
			onClose?.();
		},
	});
}
