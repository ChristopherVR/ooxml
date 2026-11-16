import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor } from 'pptx-viewer-core';

import type { CommentsPanelBinding } from '../editor/editor-comment-actions';
import type { Translator } from '../i18n';
import { createEl } from '../render';
import { renderCommentBody } from './comment-body';
import { attachCommentMentionTypeahead } from './comment-mention-typeahead';

/**
 * The threaded comment list: nested replies, edit-in-place, reply, resolve and
 * delete.
 *
 * Vanilla shows comments in TWO places (the inspector's Comments tab and the
 * workspace Comments pane) and only the inspector used to offer threads, so a
 * user who added a comment from the canvas landed in the pane with no way to
 * reply. Both surfaces now render this one view, which is also what keeps the
 * two from drifting apart. Class names stay `pptxv-inspector-comment*` because
 * that is what the stylesheet already carries.
 *
 * Replying is offered on top-level comments only (React's `depth === 0` rule);
 * editing works on replies too. The edit and reply drafts are the viewer's
 * shared comments panel controller, so they survive a re-render and a re-open.
 */

export interface CommentThreadView {
	/** The list container; append it wherever the surface wants the threads. */
	el: HTMLElement;
	/** Re-render from the current model. */
	update(comments: readonly PptxComment[], editable: boolean): void;
}

export function createCommentThreadView(
	doc: Document,
	t: Translator,
	/** The viewer's shared comments panel state: the edit and reply drafts live there. */
	panel: CommentsPanelBinding,
	/** Authors offered by the reply draft's `@`-mention typeahead. */
	getMentionAuthors: () => readonly PptxModernCommentAuthor[] = () => [],
): CommentThreadView {
	const el = createEl(doc, 'div', 'pptxv-inspector-comment-list');

	const state = () => panel.controller.getSnapshot();
	const editingId = (): string | null => state().editingIdBySlideId[panel.slideId()] ?? null;
	const replyDraft = (id: string): string => state().replyDraftById[id] ?? '';
	const replyMentions = (id: string): PptxCommentMention[] => [
		...(state().replyMentionsById[id] ?? []),
	];
	// Collapsed reply lists are this view's own; they survive store-driven re-renders.
	const collapsedReplies = new Set<string>();
	let comments: readonly PptxComment[] = [];
	let editable = false;

	const render = (): void => {
		el.replaceChildren(...comments.map((comment) => renderComment(comment, editable, 0)));
	};

	const actionBtn = (label: string, onClick: () => void): HTMLButtonElement => {
		const btn = createEl(doc, 'button', 'pptxv-inspector-comment-action');
		btn.type = 'button';
		btn.textContent = label;
		btn.addEventListener('click', onClick);
		return btn;
	};

	const draftBox = (
		value: string,
		placeholder: string,
		onInput: (next: string) => void,
	): HTMLTextAreaElement => {
		const area = doc.createElement('textarea');
		area.className = 'pptxv-inspector-comment-input';
		area.rows = 2;
		area.value = value;
		area.placeholder = placeholder;
		area.addEventListener('input', () => onInput(area.value));
		return area;
	};

	const renderEditForm = (comment: PptxComment): HTMLElement => {
		const form = createEl(doc, 'div', 'pptxv-inspector-comment-edit');
		const area = draftBox(state().editDraftById[comment.id] ?? '', '', (next) =>
			panel.controller.setEditDraft(comment.id, next),
		);
		area.setAttribute('aria-label', t('pptx.comments.edit'));
		const actions = createEl(doc, 'div', 'pptxv-inspector-comment-actions');
		actions.append(
			actionBtn(t('pptx.comments.save'), () => {
				panel.controller.saveEdit(panel.slideIndex(), comment.id);
				render();
			}),
			actionBtn(t('pptx.comments.cancel'), () => {
				panel.controller.cancelEdit(panel.slideId());
				render();
			}),
		);
		form.append(area, actions);
		return form;
	};

	const renderReplyForm = (comment: PptxComment): HTMLElement => {
		const form = createEl(doc, 'div', 'pptxv-inspector-comment-reply-form');
		const area = draftBox(
			replyDraft(comment.id),
			t('pptx.comments.replyPlaceholder', {
				author: comment.author || t('pptx.comments.unknownAuthor'),
			}),
			(next) => panel.controller.setReplyDraft(comment.id, next),
		);
		area.setAttribute('aria-label', t('pptx.comments.reply'));
		const mentionTypeahead = attachCommentMentionTypeahead({
			doc,
			t,
			field: area,
			getAuthors: getMentionAuthors,
			getMentions: () => replyMentions(comment.id),
			onChange: (next) => {
				area.value = next.text;
				panel.controller.setReplyDraft(comment.id, next.text, next.mentions);
			},
		});
		const actions = createEl(doc, 'div', 'pptxv-inspector-comment-actions');
		actions.append(
			actionBtn(t('pptx.comments.addReply'), () => {
				mentionTypeahead.destroy();
				panel.controller.submitReply(panel.slideIndex(), comment.id);
				render();
			}),
			actionBtn(t('pptx.comments.cancel'), () => {
				mentionTypeahead.destroy();
				panel.controller.cancelReply();
				render();
			}),
		);
		form.append(area, mentionTypeahead.el, actions);
		return form;
	};

	const renderActions = (comment: PptxComment, depth: number): HTMLElement => {
		const actions = createEl(doc, 'div', 'pptxv-inspector-comment-actions');
		actions.appendChild(
			actionBtn(t('pptx.comments.edit'), () => {
				panel.controller.startEdit(panel.slideId(), comment.id);
				render();
			}),
		);
		if (depth === 0) {
			actions.appendChild(
				actionBtn(t('pptx.comments.reply'), () => {
					panel.controller.startReply(comment.id);
					render();
				}),
			);
		}
		actions.append(
			actionBtn(t(comment.resolved ? 'pptx.comments.unresolve' : 'pptx.comments.resolve'), () =>
				panel.controller.toggleResolved(panel.slideIndex(), comment.id),
			),
			actionBtn(t('pptx.comments.delete'), () =>
				panel.controller.remove(panel.slideIndex(), comment.id),
			),
		);
		return actions;
	};

	const renderReplies = (
		comment: PptxComment,
		rowEditable: boolean,
		depth: number,
	): HTMLElement => {
		const wrap = createEl(doc, 'div', 'pptxv-inspector-comment-replies-wrap');
		const replies = comment.replies ?? [];
		const toggle = createEl(doc, 'button', 'pptxv-inspector-comment-replies-toggle');
		toggle.type = 'button';
		toggle.textContent = t('pptx.comments.repliesCount', { count: replies.length });
		toggle.setAttribute('aria-expanded', String(!collapsedReplies.has(comment.id)));
		toggle.addEventListener('click', () => {
			if (collapsedReplies.has(comment.id)) {
				collapsedReplies.delete(comment.id);
			} else {
				collapsedReplies.add(comment.id);
			}
			render();
		});
		wrap.appendChild(toggle);
		if (!collapsedReplies.has(comment.id)) {
			const thread = createEl(doc, 'div', 'pptxv-inspector-comment-replies');
			thread.append(...replies.map((reply) => renderComment(reply, rowEditable, depth + 1)));
			wrap.appendChild(thread);
		}
		return wrap;
	};

	function renderComment(comment: PptxComment, rowEditable: boolean, depth: number): HTMLElement {
		const row = createEl(doc, 'div', 'pptxv-inspector-comment');
		row.dataset.commentId = comment.id;
		row.classList.toggle('is-resolved', comment.resolved === true);
		row.classList.toggle('is-reply', depth > 0);
		const meta = createEl(doc, 'div', 'pptxv-inspector-comment-meta');
		meta.textContent = comment.author || t('pptx.comments.unknownAuthor');
		if (comment.resolved) {
			const badge = createEl(doc, 'span', 'pptxv-inspector-comment-badge');
			badge.textContent = t('pptx.comments.resolved');
			meta.appendChild(badge);
		}
		row.appendChild(meta);

		if (editingId() === comment.id) {
			row.appendChild(renderEditForm(comment));
		} else {
			const text = createEl(doc, 'p', 'pptxv-inspector-comment-text');
			renderCommentBody(text, comment.text, comment.mentions);
			row.appendChild(text);
			if (rowEditable) {
				row.appendChild(renderActions(comment, depth));
			}
		}

		if ((comment.replies ?? []).length > 0) {
			row.appendChild(renderReplies(comment, rowEditable, depth));
		}
		if (state().replyingToId === comment.id) {
			row.appendChild(renderReplyForm(comment));
		}
		return row;
	}

	return {
		el,
		update(nextComments, nextEditable) {
			comments = nextComments;
			editable = nextEditable;
			render();
		},
	};
}
