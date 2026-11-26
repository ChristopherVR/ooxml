import { afterEach, expect, test, vi } from 'vitest';

import { PptxViewer } from './PptxViewer';

let viewer: PptxViewer | undefined;

afterEach(() => {
	viewer?.destroy();
	viewer = undefined;
	document.body.replaceChildren();
});

test('docks comments to the viewer root when the stock body is unavailable', () => {
	const host = document.createElement('div');
	document.body.append(host);
	viewer = new PptxViewer(host);
	const body = host.querySelector('.pptxv-body');
	const root = host.querySelector('[data-pptx-editor-chrome]');
	expect(body).not.toBeNull();
	expect(root).not.toBeNull();
	body?.classList.remove('pptxv-body');

	viewer.openComments();
	expect(root?.querySelector(':scope > [data-pptx-comments-panel]')).not.toBeNull();
	viewer.openComments(true);
	expect(root?.querySelector('[data-pptx-comments-panel]')).toBeNull();
});

test('keeps the header Comments button pressed while its pane is open and returns focus on close', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	viewer = new PptxViewer(host);
	const actions = host.querySelector('pptx-ui-ribbon-actions')?.shadowRoot;
	const opener = actions?.querySelector<HTMLButtonElement>('[part~="comments"]');
	expect(opener?.getAttribute('aria-pressed')).toBe('false');

	opener?.focus();
	opener?.click();
	expect(opener?.getAttribute('aria-pressed')).toBe('true');
	const close = host.querySelector<HTMLButtonElement>('[data-pptx-comment-close]');
	expect(close).not.toBeNull();
	close?.focus();
	close?.click();
	expect(opener?.getAttribute('aria-pressed')).toBe('false');
	await vi.waitFor(() => expect(actions?.activeElement).toBe(opener));
});
