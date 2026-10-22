// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountPageNavigation,
	mountStatusBar,
	mountToolbar,
} from './building-blocks';
import { createDocument } from '../../core/src/index';

afterEach(() => document.body.replaceChildren());

it('composes a surface with library controls from the shared session and disposes cleanly', () => {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Vanilla composition here' }] }];
	const blocks = createViewerBuildingBlocks({ documentModel: model, canEdit: true });
	const part = (id: string) => {
		const node = document.createElement('div');
		node.id = id;
		document.body.append(node);
		return node;
	};
	const canvas = part('canvas');
	const toolbar = part('toolbar');
	const navigation = part('nav');
	const status = part('status');
	const stops = [
		mountDocumentCanvas(canvas, blocks),
		mountToolbar(toolbar, blocks),
		mountPageNavigation(navigation, blocks),
		mountStatusBar(status, blocks),
	];
	try {
		expect(canvas.querySelectorAll('docx-editor')).toHaveLength(1);
		expect(
			canvas.querySelectorAll('office-ui-ribbon, office-ui-title-bar, office-ui-backstage'),
		).toHaveLength(0);
		expect(toolbar.getAttribute('role')).toBe('toolbar');
		const bold = toolbar.querySelector<HTMLButtonElement>('[data-control="bold"]')!;
		expect(bold.getAttribute('aria-pressed')).toBe('false');
		bold.click();
		expect(toolbar.querySelector('[data-control="bold"]')).toBe(bold);
		expect(bold.getAttribute('aria-pressed')).toBe('true');
		expect(navigation.querySelector('[data-control="pageText"]')?.textContent).toBe('Page 1 of 1');
		expect(status.querySelector('[data-control="wordText"]')?.textContent).toBe('3 words');
		blocks.update({ documentModel: model, canEdit: false });
		expect(toolbar.querySelector<HTMLButtonElement>('[data-control="bold"]')!.disabled).toBe(true);
		expect(blocks.state.mode).toBe('view');
	} finally {
		for (const stop of stops.reverse()) stop();
		blocks.destroy();
	}
	expect(canvas.childElementCount).toBe(0);
	expect(toolbar.childElementCount).toBe(0);
});
