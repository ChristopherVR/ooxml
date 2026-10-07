import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { ViewerPageOrder } from './viewer-page-order.js';
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
			pages: ['0', '1', '2'].map((id) => ({ id, contents: `<Shapes>${shape('1')}</Shapes>` })),
		}),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const order = new ViewerPageOrder(host.attachShadow({ mode: 'open' }), controller);
	const unsubscribe = controller.subscribe((state) => order.render(state));
	const press = (name: string) =>
		order.dialog
			.querySelector(`[label="${name}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	return {
		controller,
		order,
		press,
		dispose() {
			unsubscribe();
			controller.destroy();
			host.remove();
		},
	};
}
it('cancels order drafts and applies them atomically while preserving page identity through history', async () => {
	const { controller, order, press, dispose } = await setup();
	try {
		controller.setPage(1);
		controller.selectShape({ id: '1', name: 'selected', pageId: '1' });
		order.show();
		press('Move Up');
		expect(Array.from(order.list.options, (option) => option.value)).toEqual(['1', '0', '2']);
		press('Cancel');
		expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['0', '1', '2']);
		order.show();
		press('Move Down');
		press('OK');
		await vi.waitFor(() =>
			expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['0', '2', '1']),
		);
		expect(controller.state.pageIndex).toBe(2);
		expect(controller.state.selectedShape).toMatchObject({ id: '1', pageId: '1' });
		expect(order.dialog.open).toBe(false);
		await controller.undo();
		expect(controller.state.pageIndex).toBe(1);
		await controller.redo();
		expect(controller.state.pageIndex).toBe(2);
	} finally {
		dispose();
	}
});
it('closes stale drafts on source replacement and does not apply their order to a model-only document', async () => {
	const { controller, order, press, dispose } = await setup();
	try {
		order.show();
		press('Move Down');
		const next = structuredClone(controller.state.document!);
		controller.setDocument(next);
		expect(order.dialog.open).toBe(false);
		press('OK');
		expect(controller.state.document!.pages.map((page) => page.id)).toEqual(['0', '1', '2']);
		expect(controller.state.edit.busy).toBe(false);
	} finally {
		dispose();
	}
});
