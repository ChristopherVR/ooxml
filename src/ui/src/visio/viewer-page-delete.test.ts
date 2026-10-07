import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { ViewerPageDelete } from './viewer-page-delete';
import { registerViewerControls } from './office-ui';
import { fixture, shape } from '../../../core/visio/test-fixtures';
async function setup(count = 2) {
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
			pages: Array.from({ length: count }, (_, i) => ({
				id: String(i),
				contents: `<Shapes>${shape('1')}</Shapes>`,
			})),
		}),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const remove = new ViewerPageDelete(host.attachShadow({ mode: 'open' }), controller);
	const unsubscribe = controller.subscribe((state) => remove.render(state));
	const press = (name: string) =>
		remove.dialog
			.querySelector(`[label="${name}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	return {
		controller,
		remove,
		press,
		dispose() {
			unsubscribe();
			controller.destroy();
			host.remove();
		},
	};
}
it('cancels deletion and supports undo/redo without retaining selection from a removed page', async () => {
	const { controller, remove, press, dispose } = await setup();
	try {
		controller.selectShape({ id: '1', name: 'selected', pageId: '0' });
		remove.show();
		press('Cancel');
		expect(controller.state.document!.pages).toHaveLength(2);
		remove.show();
		press('Delete');
		await vi.waitFor(() =>
			expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['1']),
		);
		expect(controller.state.selectedShape).toBeNull();
		expect(controller.state.pageIndex).toBe(0);
		expect(remove.dialog.open).toBe(false);
		await controller.undo();
		expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['0', '1']);
		await controller.redo();
		expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['1']);
	} finally {
		dispose();
	}
});
it('deletes the captured page after navigation and closes stale dialogs after source replacement', async () => {
	const { controller, remove, press, dispose } = await setup(3);
	try {
		remove.show();
		controller.setPage(2);
		press('Delete');
		await vi.waitFor(() =>
			expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['1', '2']),
		);
		expect(controller.state.pageIndex).toBe(1);
		remove.show();
		controller.setDocument(structuredClone(controller.state.document!));
		expect(remove.dialog.open).toBe(false);
		press('Delete');
		remove.show();
		expect(remove.dialog.open).toBe(false);
		expect(controller.state.document!.pages).toHaveLength(2);
	} finally {
		dispose();
	}
});
it('keeps a blank foreground page when deleting the last page', async () => {
	const { controller, remove, press, dispose } = await setup(1);
	try {
		remove.show();
		press('Delete');
		await vi.waitFor(() => expect(controller.state.document!.pages[0]!.shapes).toHaveLength(0));
		expect(controller.state.document!.pages[0]!.id).toBe('1');
		await controller.undo();
		expect(controller.state.document!.pages[0]!.shapes).toHaveLength(1);
	} finally {
		dispose();
	}
});
