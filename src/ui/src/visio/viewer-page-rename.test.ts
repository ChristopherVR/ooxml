import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { ViewerPageRename } from './viewer-page-rename.js';
import { registerViewerControls } from './office-ui.js';
import { fixture, shape } from '../../../core/visio/test-fixtures.js';

async function setup() {
	registerViewerControls();
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (source, edits) => {
			const saved = await editVsdx(source, edits);
			return { ...saved, document: await parseVsdx(saved.bytes) };
		},
	);
	await controller.load(
		await fixture({
			pages: ['0', '1'].map((id) => ({ id, contents: `<Shapes>${shape('1')}</Shapes>` })),
		}),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const rename = new ViewerPageRename(host.attachShadow({ mode: 'open' }), controller);
	const unsubscribe = controller.subscribe((state) => rename.render(state));
	const press = (name: string) =>
		rename.dialog
			.querySelector(`[label="${name}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	return {
		controller,
		rename,
		press,
		dispose() {
			unsubscribe();
			controller.destroy();
			host.remove();
		},
	};
}
it('cancels drafts and preserves page and shape identity through rename, undo and redo', async () => {
	const { controller, rename, press, dispose } = await setup();
	try {
		controller.selectShape({ id: '1', name: 'selected', pageId: '0' });
		rename.show();
		rename.input.value = 'Cancelled';
		press('Cancel');
		expect(controller.state.document!.pages[0]!.name).toBe('Page 1');
		rename.show();
		rename.input.value = 'Renamed & Page';
		press('OK');
		await vi.waitFor(() =>
			expect(controller.state.document!.pages[0]!.name).toBe('Renamed & Page'),
		);
		expect(controller.state.pageIndex).toBe(0);
		expect(controller.state.selectedShape).toMatchObject({ id: '1', pageId: '0' });
		expect(rename.dialog.open).toBe(false);
		await controller.undo();
		expect(controller.state.document!.pages[0]!.name).toBe('Page 1');
		await controller.redo();
		expect(controller.state.document!.pages[0]!.name).toBe('Renamed & Page');
	} finally {
		dispose();
	}
});
it('keeps an invalid name draft open for correction and retains its original target page', async () => {
	const { controller, rename, press, dispose } = await setup();
	try {
		rename.show();
		rename.input.value = 'Page 2';
		press('OK');
		await vi.waitFor(() => expect(rename.error.textContent).toContain('already exists'));
		expect(rename.dialog.open).toBe(true);
		controller.setPage(1);
		rename.input.value = 'Original target';
		press('OK');
		await vi.waitFor(() =>
			expect(controller.state.document!.pages[0]!.name).toBe('Original target'),
		);
		expect(controller.state.document!.pages[1]!.name).toBe('Page 2');
		expect(controller.state.pageIndex).toBe(1);
	} finally {
		dispose();
	}
});
it('closes stale drafts and does not mutate a replacement or model-only source', async () => {
	const { controller, rename, press, dispose } = await setup();
	try {
		rename.show();
		rename.input.value = 'Stale';
		controller.setDocument(structuredClone(controller.state.document!));
		expect(rename.dialog.open).toBe(false);
		press('OK');
		rename.show();
		expect(rename.dialog.open).toBe(false);
		expect(controller.state.document!.pages[0]!.name).toBe('Page 1');
	} finally {
		dispose();
	}
});
