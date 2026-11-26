// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';

import { activateCommentMarker, attachCommentsTaskPane } from './comments-task-pane';

afterEach(() => {
	document.body.replaceChildren();
	vi.useRealTimers();
});

it('keeps focus on a comment selected just after opening the pane', async () => {
	vi.useFakeTimers();
	const pane = document.createElement('aside');
	const row = document.createElement('div');
	row.tabIndex = -1;
	pane.append(row);
	document.body.append(pane);
	const detach = attachCommentsTaskPane(pane, {
		label: 'Comments',
		closeLabel: 'Close',
		close: () => undefined,
	});
	row.focus();
	await vi.advanceTimersByTimeAsync(30);
	expect(document.activeElement).toBe(row);
	detach();
});

it('opens the thread named by the marker and moves to a second thread in an open pane', async () => {
	vi.useFakeTimers();
	const root = document.createElement('div');
	root.dataset.pptxEditorChrome = '';
	const pane = document.createElement('aside');
	const draft = document.createElement('textarea');
	draft.dataset.pptxCommentDraft = '';
	pane.append(draft);
	const thread = (id: string) => {
		const row = document.createElement('article');
		row.dataset.commentId = id;
		pane.append(row);
		const marker = document.createElement('button');
		marker.dataset.pptxCommentMarker = id;
		root.append(marker);
		return { row, marker };
	};
	const first = thread('first');
	const second = thread('second');
	document.body.append(root);
	root.append(pane);
	let detach: (() => void) | undefined;
	first.marker.addEventListener('click', (event) =>
		activateCommentMarker(event, () => {
			detach = attachCommentsTaskPane(pane, {
				label: 'Comments',
				closeLabel: 'Close',
				close: () => undefined,
			});
		}),
	);
	second.marker.addEventListener('click', (event) => activateCommentMarker(event, () => {}));
	first.marker.click();
	await vi.advanceTimersByTimeAsync(30);
	expect(document.activeElement).toBe(first.row);
	second.marker.click();
	expect(document.activeElement).toBe(second.row);
	expect(first.row.hasAttribute('data-comment-active')).toBe(false);
	detach?.();
});

it('focuses the draft, closes from Escape, and restores the opener inside a shadow root', async () => {
	vi.useFakeTimers();
	const host = document.createElement('div');
	const opener = document.createElement('button');
	host.attachShadow({ mode: 'open' }).append(opener);
	const pane = document.createElement('aside');
	const draft = document.createElement('textarea');
	draft.setAttribute('data-pptx-comment-draft', '');
	pane.append(draft);
	document.body.append(host, pane);
	opener.focus();
	const detach = attachCommentsTaskPane(pane, {
		label: 'Comments',
		closeLabel: 'Close',
		close: () => {
			detach();
			pane.remove();
		},
	});
	await vi.advanceTimersByTimeAsync(30);
	expect(document.activeElement).toBe(draft);
	expect(pane.querySelector('button')?.getAttribute('aria-label')).toBe('Close');
	draft.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	await vi.advanceTimersByTimeAsync(0);
	expect(pane.isConnected).toBe(false);
	expect(host.shadowRoot?.activeElement).toBe(opener);
});
