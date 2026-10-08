import { expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { ViewerController } from './controller';

const targets = ['s1', 's2'].map((id) => ({ id, name: id, pageId: '1' }));
function setup(remove = true) {
	const original = structuredClone(demoDocument);
	const changed = structuredClone(original);
	if (remove)
		changed.pages[0]!.shapes = changed.pages[0]!.shapes.filter(
			(shape) => !targets.some((target) => target.id === shape.id),
		);
	const parser = vi.fn(async (bytes: Uint8Array | ArrayBuffer) =>
		new Uint8Array(bytes instanceof Uint8Array ? bytes : bytes)[0] === 1 ? original : changed,
	);
	const controller = new ViewerController(
		parser,
		() => {},
		async () => ({
			bytes: new Uint8Array([2]),
			document: changed,
			changedParts: ['page.xml'],
			diagnostics: [],
		}),
	);
	return { controller, parser, original };
}
async function deleted() {
	const ui = setup();
	await ui.controller.load(new Uint8Array([1]));
	ui.controller.selectShapes(targets);
	await ui.controller.applyEdits(
		targets.map((target) => ({ type: 'delete-shape', pageId: '1', shapeId: target.id })),
	);
	expect(ui.controller.state.selectedShapes).toEqual([]);
	return ui;
}
function defer<T>() {
	let resolve!: (value: T) => void;
	const promise = new Promise<T>((done) => {
		resolve = done;
	});
	return { promise, resolve };
}

it('restores all deleted targets in their primary order on undo and clears them again on redo', async () => {
	const { controller } = await deleted();
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual(targets);
	expect(controller.state.selectedShape).toBe(controller.state.selectedShapes[0]);
	expect(Object.isFrozen(controller.state.selectedShapes)).toBe(true);
	expect(controller.state.selectedShapes.every(Object.isFrozen)).toBe(true);
	await controller.redo();
	expect(controller.state.selectedShapes).toEqual([]);
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual(targets);
	controller.destroy();
});

it.each(['select', 'clear'] as const)(
	'honors explicit user %s intent after deletion, including clearing an already empty selection',
	async (mode) => {
		const { controller } = await deleted();
		const chosen = [{ id: 's4', name: 'Other', pageId: '1' }];
		if (mode === 'clear') controller.clearSelection();
		else controller.selectShapes(chosen);
		await controller.undo();
		expect(controller.state.selectedShapes).toEqual(mode === 'clear' ? [] : chosen);
		await controller.redo();
		expect(controller.state.selectedShapes).toEqual(mode === 'clear' ? [] : chosen);
		controller.destroy();
	},
);

it.each(['select', 'clear', 'page'] as const)(
	'honors newer %s intent while an undo source is being parsed',
	async (mode) => {
		const { controller, parser, original } = await deleted();
		const waiting = defer<typeof original>();
		parser.mockImplementationOnce(() => waiting.promise);
		const operation = controller.undo();
		const chosen = [{ id: 's4', name: 'Other', pageId: '1' }];
		if (mode === 'select') controller.selectShapes(chosen);
		else if (mode === 'clear') controller.clearSelection();
		else controller.setPage(1);
		waiting.resolve(original);
		await operation;
		expect(controller.state.selectedShapes).toEqual(mode === 'select' ? chosen : []);
		expect(controller.state.pageIndex).toBe(mode === 'page' ? 1 : 0);
		controller.destroy();
	},
);

it('preserves the current selection during history edits that did not prune it', async () => {
	const { controller } = setup(false);
	await controller.load(new Uint8Array([1]));
	controller.selectShapes([targets[0]!]);
	await controller.applyEdits([{ type: 'format-text', pageId: '1', shapeId: 's1', bold: true }]);
	controller.selectShapes([targets[1]!]);
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual([targets[1]!]);
	await controller.redo();
	expect(controller.state.selectedShapes).toEqual([targets[1]!]);
	controller.destroy();
});

it('cancels pending restoration without changing the accepted source or losing later undo metadata', async () => {
	const { controller, parser, original } = await deleted();
	const waiting = defer<typeof original>();
	parser.mockImplementationOnce(() => waiting.promise);
	const operation = controller.undo();
	const rejected = expect(operation).rejects.toHaveProperty('name', 'AbortError');
	controller.cancelEdit();
	waiting.resolve(original);
	await rejected;
	expect(controller.exportVsdx().bytes).toEqual(new Uint8Array([2]));
	expect(controller.state.selectedShapes).toEqual([]);
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual(targets);
	controller.destroy();
});

it('drops old source selection metadata when a pending undo is superseded by replacement', async () => {
	const { controller, parser, original } = await deleted();
	const waiting = defer<typeof original>();
	parser.mockImplementationOnce(() => waiting.promise);
	const operation = controller.undo();
	const rejected = expect(operation).rejects.toHaveProperty('name', 'AbortError');
	controller.setDocument(demoDocument);
	waiting.resolve(original);
	await rejected;
	expect(controller.state.selectedShapes).toEqual([]);
	expect(controller.state.edit.canUndo).toBe(false);
	await controller.load(new Uint8Array([1]));
	expect(controller.state.selectedShapes).toEqual([]);
	expect(controller.state.edit.canUndo).toBe(false);
	controller.destroy();
});

it('does not reuse discarded redo selection metadata after history branching', async () => {
	const { controller } = await deleted();
	await controller.undo();
	controller.clearSelection();
	await controller.applyEdits([{ type: 'format-text', pageId: '1', shapeId: 's4', bold: true }]);
	expect(controller.state.edit.canRedo).toBe(false);
	await controller.undo();
	expect(controller.state.selectedShapes).toEqual([]);
	controller.destroy();
});
