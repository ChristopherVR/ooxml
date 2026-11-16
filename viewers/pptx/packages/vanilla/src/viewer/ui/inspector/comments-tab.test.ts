import type { PptxComment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../../i18n';
import { createTestCommentsPanel } from '../comments-panel.test-support';
import { createCommentsTab } from './comments-tab';
import type { InspectorDeckState } from './types';

const t = createTranslator();

/** A tab over the shared controller; `refresh` re-renders from the deck after a command. */
function openTab(
	comments: PptxComment[],
	editable = true,
	extra: Partial<InspectorDeckState> = {},
) {
	const deck = createTestCommentsPanel(comments, editable);
	const tab = createCommentsTab(document, t, { commentsPanel: () => deck.panel });
	const refresh = (): void =>
		tab.update({ comments: deck.comments(), editable, ...extra } as unknown as InspectorDeckState);
	refresh();
	return { tab, deck, refresh };
}

function findAction(scope: Element, label: string): HTMLButtonElement {
	const btn = Array.from(
		scope.querySelectorAll<HTMLButtonElement>('.pptxv-inspector-comment-action'),
	).find((b) => b.textContent === label);
	expect(btn).toBeDefined();
	return btn!;
}

const ALICE = { id: 'a1', name: 'Alice', userId: '', providerId: '' };

/** Type `value` into `field` with the caret at its end, then pick the first suggestion. */
function typeAndPickMention(scope: Element, field: HTMLTextAreaElement, value: string): void {
	field.value = value;
	field.selectionStart = value.length;
	field.selectionEnd = value.length;
	field.dispatchEvent(new Event('input'));
	scope
		.querySelector<HTMLButtonElement>('[data-testid="pptx-comment-mention-option"]')!
		.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
}

const covered = ({ text, mentions }: PptxComment) =>
	(mentions ?? []).map((m) => text!.slice(m.startIndex, m.startIndex + m.length));

describe('inspector comments tab', () => {
	it('adds a comment and a reply with their mentions on the trimmed text', () => {
		const { tab, deck, refresh } = openTab([{ id: 'c1', text: 'Parent', author: 'Ada' }], true, {
			commentMentionAuthors: [ALICE],
		});
		const input = tab.el.querySelector<HTMLTextAreaElement>('.pptxv-inspector-comment-input')!;
		typeAndPickMention(tab.el, input, '  hi @al');
		tab.el.querySelector<HTMLButtonElement>('.pptxv-inspector-deck-btn')!.click();
		const added = deck.comments().at(-1)!;
		expect(added.text).toBe('hi @Alice');
		expect(covered(added)).toStrictEqual(['@Alice']);
		refresh();

		findAction(tab.el, t('pptx.comments.reply')).click();
		const form = tab.el.querySelector<HTMLElement>('.pptxv-inspector-comment-reply-form')!;
		typeAndPickMention(form, form.querySelector('textarea')!, ' ok @al');
		findAction(form, t('pptx.comments.addReply')).click();
		const reply = deck.comments()[0]!.replies![0]!;
		expect(reply.text).toBe('ok @Alice');
		expect(covered(reply)).toStrictEqual(['@Alice']);
	});

	it('keeps the new-comment draft with its slide', () => {
		const { tab, deck, refresh } = openTab([]);
		const input = tab.el.querySelector<HTMLTextAreaElement>('.pptxv-inspector-comment-input')!;
		input.value = 'For slide one';
		input.dispatchEvent(new Event('input'));
		deck.setCurrent(1);
		refresh();
		expect(input.value).toBe('');
		deck.setCurrent(0);
		refresh();
		expect(input.value).toBe('For slide one');
	});

	it('renders nested replies with a count toggle', () => {
		const { tab } = openTab([
			{
				id: 'c1',
				text: 'Parent',
				author: 'Ada',
				replies: [
					{ id: 'r1', text: 'First reply', author: 'Bob', threadId: 'c1' },
					{ id: 'r2', text: 'Second reply', author: 'Cee', threadId: 'c1' },
				],
			},
		]);
		const replies = tab.el.querySelectorAll('.pptxv-inspector-comment.is-reply');
		expect(replies).toHaveLength(2);
		const toggle = tab.el.querySelector<HTMLButtonElement>(
			'.pptxv-inspector-comment-replies-toggle',
		);
		expect(toggle?.textContent).toBe(t('pptx.comments.repliesCount', { count: 2 }));
		toggle!.click();
		expect(tab.el.querySelectorAll('.pptxv-inspector-comment.is-reply')).toHaveLength(0);
		toggle!.click();
		expect(tab.el.querySelectorAll('.pptxv-inspector-comment.is-reply')).toHaveLength(2);
	});

	it('submits a reply and closes the form', () => {
		const { tab, deck } = openTab([{ id: 'c1', text: 'Parent', author: 'Ada' }]);
		findAction(tab.el, t('pptx.comments.reply')).click();
		const form = tab.el.querySelector<HTMLElement>('.pptxv-inspector-comment-reply-form');
		const area = form!.querySelector<HTMLTextAreaElement>('textarea');
		area!.value = 'A fresh reply';
		area!.dispatchEvent(new Event('input'));
		findAction(form!, t('pptx.comments.addReply')).click();
		expect(deck.comments()[0]!.replies![0]!.text).toBe('A fresh reply');
		expect(tab.el.querySelector('.pptxv-inspector-comment-reply-form')).toBeNull();
	});

	it('does not offer reply on nested reply rows (depth 0 only)', () => {
		const { tab } = openTab([
			{ id: 'c1', text: 'Parent', replies: [{ id: 'r1', text: 'Reply', threadId: 'c1' }] },
		]);
		const replyRow = tab.el.querySelector('.pptxv-inspector-comment.is-reply');
		const labels = Array.from(
			replyRow!.querySelectorAll<HTMLButtonElement>('.pptxv-inspector-comment-action'),
		).map((b) => b.textContent);
		expect(labels).not.toContain(t('pptx.comments.reply'));
		expect(labels).toContain(t('pptx.comments.edit'));
	});

	it('edits a comment, and a nested reply, in place', () => {
		const { tab, deck, refresh } = openTab([
			{ id: 'c1', text: 'Original', replies: [{ id: 'r1', text: 'Reply', threadId: 'c1' }] },
		]);
		findAction(tab.el, t('pptx.comments.edit')).click();
		let editBox = tab.el.querySelector<HTMLElement>('.pptxv-inspector-comment-edit')!;
		let area = editBox.querySelector<HTMLTextAreaElement>('textarea')!;
		expect(area.value).toBe('Original');
		area.value = 'Rewritten';
		area.dispatchEvent(new Event('input'));
		findAction(editBox, t('pptx.comments.save')).click();
		expect(deck.comments()[0]!.text).toBe('Rewritten');
		expect(tab.el.querySelector('.pptxv-inspector-comment-edit')).toBeNull();
		refresh();

		const replyRow = tab.el.querySelector('.pptxv-inspector-comment.is-reply');
		findAction(replyRow!, t('pptx.comments.edit')).click();
		editBox = tab.el.querySelector<HTMLElement>('.pptxv-inspector-comment-edit')!;
		area = editBox.querySelector<HTMLTextAreaElement>('textarea')!;
		area.value = 'Reply v2';
		area.dispatchEvent(new Event('input'));
		findAction(editBox, t('pptx.comments.save')).click();
		expect(deck.comments()[0]!.replies![0]!.text).toBe('Reply v2');
	});

	it('cancelling an edit leaves the comment untouched', () => {
		const { tab, deck } = openTab([{ id: 'c1', text: 'Original' }]);
		findAction(tab.el, t('pptx.comments.edit')).click();
		const editBox = tab.el.querySelector<HTMLElement>('.pptxv-inspector-comment-edit');
		findAction(editBox!, t('pptx.comments.cancel')).click();
		expect(deck.comments()[0]!.text).toBe('Original');
		expect(tab.el.querySelector('.pptxv-inspector-comment-edit')).toBeNull();
	});

	it('hides editing affordances in read-only mode but still shows replies', () => {
		const { tab } = openTab(
			[{ id: 'c1', text: 'Parent', replies: [{ id: 'r1', text: 'Reply', threadId: 'c1' }] }],
			false,
		);
		expect(tab.el.querySelectorAll('.pptxv-inspector-comment-action')).toHaveLength(0);
		expect(tab.el.querySelectorAll('.pptxv-inspector-comment.is-reply')).toHaveLength(1);
	});
});
