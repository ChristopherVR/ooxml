import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { ViewerController } from './controller';
import type { CancellableEditor } from './worker-editor';
const edit: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};
async function setup(editor: CancellableEditor = edit) {
	const controller = new ViewerController(parseVsdx, () => {}, editor);
	await controller.load(await createVsdxFixture());
	controller.selectAll();
	return controller;
}
const move = { type: 'move-shape' as const, pageId: '1', shapeId: '1', x: 2, y: 3 };
it('accepts a selected shape transaction with undo while refusing unrelated targets', async () => {
	const controller = await setup();
	const before = controller.exportVsdx().bytes;
	await expect(controller.applySelectionEdits([{ ...move, shapeId: '2' }])).rejects.toThrow(
		'selected shapes',
	);
	await expect(
		controller.applySelectionEdits([
			{ type: 'insert-page', pageId: '2', name: 'Other', afterPageId: '1' },
		]),
	).rejects.toThrow('selected shapes');
	await controller.applySelectionEdits([move]);
	expect(controller.state.selectedShape?.id).toBe('1');
	expect(controller.state.document!.pages[0]!.shapes[0]).toMatchObject({
		rotation: { pinX: 2, pinY: 3 },
	});
	await controller.undo();
	expect(controller.exportVsdx().bytes).toEqual(before);
	controller.destroy();
});
it('refuses host command getters that supersede the captured selection', async () => {
	const controller = await setup();
	const bytes = controller.exportVsdx().bytes;
	await expect(
		controller.applySelectionEdits([
			{
				...move,
				get x() {
					controller.clearSelection();
					return 2;
				},
			},
		]),
	).rejects.toHaveProperty('name', 'AbortError');
	expect(controller.exportVsdx().bytes).toEqual(bytes);
	controller.destroy();
});
it('refuses newer selection intent during an asynchronous gesture edit', async () => {
	let release!: () => void;
	const waiting = new Promise<void>((done) => {
		release = done;
	});
	const service = vi.fn(async (bytes: Uint8Array, commands: Parameters<CancellableEditor>[1]) => {
		await waiting;
		return edit(bytes, commands);
	});
	const controller = await setup(service);
	const before = controller.exportVsdx().bytes;
	const pending = controller.applySelectionEdits([move]);
	const refused = expect(pending).rejects.toHaveProperty('name', 'AbortError');
	controller.clearSelection();
	release();
	await refused;
	expect(controller.exportVsdx().bytes).toEqual(before);
	expect(controller.state.edit).toMatchObject({ canUndo: false, error: null });
	controller.destroy();
});
