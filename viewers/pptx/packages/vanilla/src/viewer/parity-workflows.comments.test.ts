import { afterEach, expect, it } from 'vitest';

import { PptxViewer } from './PptxViewer';

let viewer: PptxViewer | undefined;

afterEach(() => {
	viewer?.destroy();
	viewer = undefined;
	document.body.replaceChildren();
});

it('docks comments to the viewer root when the stock body is unavailable', () => {
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
