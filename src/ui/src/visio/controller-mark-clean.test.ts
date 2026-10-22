import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { ViewerController } from './controller';
import type { CancellableEditor } from './worker-editor';

const edit: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};
const rectangle = (shapeId: string) => ({
	type: 'create-rectangle' as const,
	pageId: '1',
	shapeId,
	x: 2,
	y: 2,
	width: 1,
	height: 1,
});
async function setup() {
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		edit,
		async () => 'unused',
	);
	await controller.load(await createVsdxFixture('Saved'));
	return controller;
}

it('clears dirty after a host save, emits the change and measures later edits from there', async () => {
	const controller = await setup();
	const changes: Array<{ dirty: boolean; kind: string }> = [];
	controller.onEvent((name, detail) => {
		if (name === 'document-change') changes.push(detail as { dirty: boolean; kind: string });
	});
	await controller.applyEdits([rectangle('20')]);
	expect(controller.state.edit.dirty).toBe(true);
	const saved = controller.exportVsdx();
	expect(saved.dirty).toBe(true);
	// exportVsdx never changes dirty: the host acknowledges the save itself.
	expect(controller.state.edit.dirty).toBe(true);
	controller.markClean();
	expect(controller.state.edit.dirty).toBe(false);
	expect(controller.exportVsdx().dirty).toBe(false);
	expect(changes.at(-1)).toMatchObject({ dirty: false, kind: 'save' });
	await controller.applyEdits([rectangle('21')]);
	expect(controller.state.edit.dirty).toBe(true);
	// Undo returns to the saved point, which is clean, not to the original.
	await controller.undo();
	expect(controller.state.edit.dirty).toBe(false);
	await controller.undo();
	expect(controller.state.edit.dirty).toBe(true);
	controller.destroy();
});

it('does not announce a save when nothing was dirty and ignores drawings without a source', async () => {
	const controller = await setup();
	const listener = vi.fn();
	controller.onEvent(listener);
	controller.markClean();
	expect(listener).not.toHaveBeenCalled();
	controller.setDocument(null);
	expect(() => controller.markClean()).not.toThrow();
	controller.destroy();
});
