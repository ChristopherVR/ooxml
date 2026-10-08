// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { defineCommentsPane, type OfficeUiCommentsPane } from './comments-pane';
import { commentInitials, formatCommentTime, type OfficeCommentThread } from './types';

beforeAll(() => defineCommentsPane());
afterEach(() => document.body.replaceChildren());

const THREADS: OfficeCommentThread[] = [
	{
		id: 't1',
		anchorLabel: 'A1',
		resolved: false,
		comments: [
			{ id: 'c1', author: 'Ada Lovelace', created: '2024-03-01T10:00:00Z', text: 'Check @Grace' },
			{ id: 'c2', author: 'Grace', initials: 'GH', text: 'Done', edited: true },
		],
	},
	{ id: 't2', resolved: true, comments: [{ id: 'c3', author: 'Ann', text: '<b>raw</b>' }] },
	{ id: 't3', resolved: false, comments: [{ id: 'c4', author: 'Bob', text: 'Third' }] },
];

function pane(setup: (el: OfficeUiCommentsPane) => void = () => {}): OfficeUiCommentsPane {
	const el = document.createElement('office-ui-comments-pane') as OfficeUiCommentsPane;
	el.threads = THREADS;
	setup(el);
	document.body.append(el);
	return el;
}
const root = (el: OfficeUiCommentsPane) => el.shadowRoot!;
const cards = (el: OfficeUiCommentsPane) => [...root(el).querySelectorAll<HTMLElement>('.thread')];
const button = (scope: ParentNode, label: string) =>
	[...scope.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === label)!;
function listen(el: HTMLElement, type: string) {
	const spy = vi.fn();
	el.addEventListener(type, (event) => spy((event as CustomEvent).detail));
	return spy;
}
function key(target: HTMLElement, name: string) {
	const event = new KeyboardEvent('keydown', {
		key: name,
		bubbles: true,
		composed: true,
		cancelable: true,
	});
	target.dispatchEvent(event);
	return event;
}

describe('office-ui-comments-pane rendering', () => {
	it('draws threads with avatar, author, time, text and replies, keeping text as text', () => {
		const el = pane((p) => (p.locale = 'en-US'));
		expect(cards(el).map((card) => card.dataset.threadId)).toEqual(['t1', 't2', 't3']);
		const [first, second] = cards(el);
		expect(first!.getAttribute('role')).toBe('listitem');
		expect(root(el).querySelector('.threads')!.getAttribute('role')).toBe('list');
		expect(first!.querySelector('.anchor-label')!.textContent).toBe('A1');
		const comments = [...first!.querySelectorAll<HTMLElement>('.comment')];
		expect(comments.map((c) => c.querySelector('.avatar')!.textContent)).toEqual(['AL', 'GH']);
		expect(comments[0]!.querySelector('.text')!.textContent).toBe('Check @Grace');
		expect(comments[0]!.querySelector('time')!.getAttribute('datetime')).toBe(
			'2024-03-01T10:00:00Z',
		);
		expect(comments[1]!.classList.contains('reply')).toBe(true);
		expect(comments[1]!.querySelector('.edited')!.textContent).toBe('Edited');
		expect(second!.querySelector('.text')!.innerHTML).not.toContain('<b>');
		expect(second!.dataset.resolved).toBe('true');
		expect(second!.querySelector('.badge')!.textContent).toBe('Resolved');
		expect(button(second!, 'Reopen')).toBeTruthy();
		expect(button(first!, 'Resolve')).toBeTruthy();
	});

	it('takes every visible string from labels and shows the empty text', () => {
		const el = pane((p) => {
			p.threads = [];
			p.closable = true;
			p.labels = { heading: 'Kommentare', empty: 'Keine', add: 'Hinzu', close: 'Schließen' };
		});
		expect(root(el).querySelector('h2')!.textContent).toBe('Kommentare');
		expect(root(el).querySelector('.empty')!.textContent).toBe('Keine');
		expect(button(root(el), 'Hinzu')).toBeTruthy();
		expect(root(el).querySelector('.close')!.getAttribute('aria-label')).toBe('Schließen');
		el.labels = { heading: '' };
		expect(root(el).querySelector('h2')).toBeNull();
	});

	it('adds product hook classes and marks the active thread', () => {
		const el = pane((p) => {
			p.classNames = { thread: 'x-thread', comment: 'x-comment' };
			p.activeThreadId = 't3';
		});
		expect(root(el).querySelectorAll('.x-thread')).toHaveLength(3);
		expect(root(el).querySelectorAll('.x-comment')).toHaveLength(4);
		expect(cards(el)[2]!.getAttribute('aria-current')).toBe('true');
		expect(cards(el).map((card) => card.tabIndex)).toEqual([-1, -1, 0]);
	});

	it('shows disabled controls when read-only and none in list-only mode', () => {
		const el = pane((p) => (p.readOnly = true));
		const controls = [...root(el).querySelectorAll<HTMLButtonElement>('button, textarea')];
		expect(controls.length).toBeGreaterThan(0);
		expect(controls.every((control) => control.disabled)).toBe(true);
		el.readOnly = false;
		el.listOnly = true;
		expect(root(el).querySelectorAll('button, textarea')).toHaveLength(0);
		expect(el.hasAttribute('list-only')).toBe(true);
	});
});

describe('office-ui-comments-pane events', () => {
	it('posts new comments and replies trimmed, ignoring blank text', () => {
		const el = pane();
		const add = listen(el, 'comment-add');
		const reply = listen(el, 'comment-reply');
		const box = root(el).querySelector<HTMLTextAreaElement>('textarea.new')!;
		button(root(el), 'Add comment').click();
		box.value = '  New one  ';
		button(root(el), 'Add comment').click();
		expect(add.mock.calls).toEqual([[{ text: 'New one' }]]);
		expect(box.value).toBe('');
		const card = cards(el)[1]!;
		card.querySelector<HTMLTextAreaElement>('textarea')!.value = 'Answer';
		button(card, 'Reply').click();
		expect(reply).toHaveBeenCalledWith({ threadId: 't2', text: 'Answer' });
	});

	it('posts with Ctrl+Enter from a box', () => {
		const el = pane();
		const add = listen(el, 'comment-add');
		const box = root(el).querySelector<HTMLTextAreaElement>('textarea.new')!;
		box.value = 'Quick';
		box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }));
		expect(add).toHaveBeenCalledWith({ text: 'Quick' });
	});

	it('fires resolve, reopen, delete and close', () => {
		const el = pane((p) => (p.closable = true));
		const resolve = listen(el, 'thread-resolve');
		const reopen = listen(el, 'thread-reopen');
		const remove = listen(el, 'comment-delete');
		const close = listen(el, 'comments-close');
		button(cards(el)[0]!, 'Resolve').click();
		button(cards(el)[1]!, 'Reopen').click();
		cards(el)[0]!.querySelectorAll<HTMLButtonElement>('[data-action="delete"]')[1]!.click();
		root(el).querySelector<HTMLButtonElement>('.close')!.click();
		expect(resolve).toHaveBeenCalledWith({ threadId: 't1' });
		expect(reopen).toHaveBeenCalledWith({ threadId: 't2' });
		expect(remove).toHaveBeenCalledWith({ threadId: 't1', commentId: 'c2' });
		expect(close).toHaveBeenCalledOnce();
	});

	it('edits only the current author comments, and Escape cancels the edit', () => {
		const el = pane((p) => (p.currentAuthor = 'Grace'));
		const edit = listen(el, 'comment-edit');
		const outer = vi.fn();
		document.body.addEventListener('keydown', outer);
		expect(root(el).querySelectorAll('[data-action="edit"]')).toHaveLength(1);
		button(root(el), 'Edit').click();
		let box = root(el).querySelector<HTMLTextAreaElement>('textarea.edit-box')!;
		expect(box.value).toBe('Done');
		key(box, 'Escape');
		expect(outer).not.toHaveBeenCalled();
		expect(root(el).querySelector('textarea.edit-box')).toBeNull();
		button(root(el), 'Edit').click();
		box = root(el).querySelector<HTMLTextAreaElement>('textarea.edit-box')!;
		box.value = 'Done twice';
		button(root(el), 'Save').click();
		expect(edit).toHaveBeenCalledWith({ threadId: 't1', commentId: 'c2', text: 'Done twice' });
		document.body.removeEventListener('keydown', outer);
	});

	it('selects a thread on click, but not from its controls', () => {
		const el = pane();
		const select = listen(el, 'thread-select');
		cards(el)[2]!.querySelector<HTMLElement>('.text')!.click();
		button(cards(el)[0]!, 'Resolve').click();
		expect(select.mock.calls).toEqual([[{ threadId: 't3' }]]);
	});
});

describe('office-ui-comments-pane keyboard', () => {
	it('moves between threads with arrows, Home and End and selects with Enter or Space', () => {
		const el = pane();
		const select = listen(el, 'thread-select');
		el.focus();
		expect(el.shadowRoot!.activeElement).toBe(cards(el)[0]);
		expect(key(cards(el)[0]!, 'ArrowDown').defaultPrevented).toBe(true);
		expect(el.shadowRoot!.activeElement).toBe(cards(el)[1]);
		expect(cards(el).map((card) => card.tabIndex)).toEqual([-1, 0, -1]);
		key(cards(el)[1]!, 'End');
		expect(el.shadowRoot!.activeElement).toBe(cards(el)[2]);
		key(cards(el)[2]!, 'ArrowDown');
		expect(el.shadowRoot!.activeElement).toBe(cards(el)[2]);
		key(cards(el)[2]!, 'Home');
		key(cards(el)[0]!, 'ArrowUp');
		expect(el.shadowRoot!.activeElement).toBe(cards(el)[0]);
		key(cards(el)[0]!, 'Enter');
		key(cards(el)[0]!, ' ');
		expect(select.mock.calls).toEqual([[{ threadId: 't1' }], [{ threadId: 't1' }]]);
	});

	it('leaves keys typed into a box alone', () => {
		const el = pane();
		const box = cards(el)[0]!.querySelector<HTMLTextAreaElement>('textarea')!;
		expect(key(box, 'ArrowDown').defaultPrevented).toBe(false);
		expect(key(box, 'Escape').defaultPrevented).toBe(false);
	});
});

describe('comment helpers', () => {
	it('derives initials and formats ISO times only', () => {
		expect(commentInitials('Ada Lovelace')).toBe('AL');
		expect(commentInitials(' grace ')).toBe('G');
		expect(commentInitials('Jean de la Fontaine')).toBe('JF');
		expect(formatCommentTime(undefined)).toBe('');
		expect(formatCommentTime('yesterday')).toBe('yesterday');
		expect(formatCommentTime('2024-03-01T10:00:00Z', 'en-US')).toMatch(/2024/);
	});
});
