import { expect, it, vi } from 'vitest';
import {
	captureVisioClipboard,
	createVsdx,
	editVsdx,
	parseVsdx,
	serializeVisioClipboard,
	type VisioEdit,
} from 'ooxml-core/visio';
import { ViewerController } from './controller';
import type { ViewerCreationToken } from './creation-token';
import type { CancellableEditor } from './worker-editor';

const create = {
	type: 'create-rectangle' as const,
	pageId: '0',
	shapeId: '2',
	x: 4,
	y: 3,
	width: 2,
	height: 1,
};
const parse = async (bytes: Uint8Array | ArrayBuffer) => {
	const model = await parseVsdx(bytes);
	// A saved source is retained independently from this host's display layer metadata.
	model.pages[0]!.layers = [
		{ id: '0', name: 'View layer', visible: true, printable: true, locked: false },
	];
	return model;
};
const edit: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parse(result.bytes) };
};
async function setup(service: CancellableEditor = edit) {
	const shaped = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 1, height: 1 },
	]);
	const source = await editVsdx(shaped.bytes, [
		{ type: 'insert-page', pageId: '1', afterPageId: '0', name: 'Other' },
	]);
	const controller = new ViewerController(
		parse,
		() => {},
		service,
		async (bytes, pageId, ids) =>
			serializeVisioClipboard(await captureVisioClipboard(bytes, pageId, ids)),
	);
	await controller.load(source.bytes);
	controller.selectShape({
		id: '1',
		name: controller.state.document!.pages[0]!.shapes[0]!.name,
		pageId: '0',
	});
	return controller;
}
function deferred() {
	let release!: () => void;
	const waiting = new Promise<void>((resolve) => {
		release = resolve;
	});
	return { waiting, release };
}
type Change = 'selection' | 'page' | 'zoom' | 'layers' | 'cancel' | 'load' | 'model';
async function change(controller: ViewerController, kind: Change, replacement?: Uint8Array) {
	if (kind === 'selection') {
		controller.clearSelection();
		controller.selectAll();
	}
	if (kind === 'page') {
		controller.setPage(1);
		controller.setPage(0);
	}
	if (kind === 'zoom') {
		controller.setZoom(2);
		controller.setZoom(1);
	}
	if (kind === 'layers') {
		controller.setLayerVisibility('0', '0', false);
		controller.resetLayerVisibility();
	}
	if (kind === 'cancel') controller.cancelEdit();
	if (kind === 'load') await controller.load(replacement ?? controller.exportVsdx().bytes);
	if (kind === 'model') controller.setDocument(controller.state.document);
}
it('accepts created source and ordered frozen primary selection in one history transition', async () => {
	const controller = await setup();
	const before = controller.exportVsdx().bytes;
	const originals = controller.state.selectedShapes;
	const token = controller.captureCreationToken('0');
	expect(Object.isFrozen(token)).toBe(true);
	expect(Object.keys(token)).toEqual([]);
	const primary = vi.fn();
	controller.onEvent((name, detail) => {
		if (name === 'shape-select') primary(detail);
	});
	const observed = vi.fn();
	controller.subscribe((state) => {
		if (state.document?.pages[0]!.shapes.length === 3 && !state.edit.busy) {
			expect(state.selectedShapes.map((item) => item.id)).toEqual(['3', '2']);
			observed();
		}
	});
	await controller.applyCreationEdits(
		[{ ...create, shapeId: '3', type: 'create-ellipse' }, create],
		token,
	);
	const created = controller.state.selectedShapes;
	expect(created.map((item) => item.id)).toEqual(['3', '2']);
	expect(Object.isFrozen(created) && created.every(Object.isFrozen)).toBe(true);
	expect(controller.state.selectedShape).toBe(created[0]);
	expect(primary).toHaveBeenCalledOnce();
	expect(observed).toHaveBeenCalled();
	expect(controller.isCreationTokenCurrent(token)).toBe(false);
	await controller.undo();
	expect(controller.exportVsdx().bytes).toEqual(before);
	expect(controller.state.selectedShapes).toEqual(originals);
	await controller.redo();
	expect(controller.state.selectedShapes).toEqual(created);
	controller.destroy();
});
it('keeps a captured intent valid through unrelated search and clipboard readiness publications', async () => {
	const controller = await setup();
	const token = controller.captureCreationToken('0');
	controller.setSearchQuery('Unrelated');
	await controller.prepareClipboardSelection(controller.captureClipboardToken());
	expect(controller.isCreationTokenCurrent(token)).toBe(true);
	await controller.applyCreationEdits([create], token);
	expect(controller.state.selectedShape?.id).toBe('2');
	controller.destroy();
});
it('creates retained multiline text with selected source and exact undo/redo history', async () => {
	const controller = await setup();
	const before = controller.exportVsdx().bytes;
	const text = 'Created text\nSecond paragraph\n';
	await controller.applyCreationEdits(
		[{ ...create, type: 'create-text-box', text }],
		controller.captureCreationToken('0'),
	);
	const created = controller.exportVsdx().bytes;
	expect(controller.state.selectedShape?.id).toBe('2');
	expect((await parseVsdx(created)).pages[0]!.shapes[1]!.text.plainText).toBe(text);
	await controller.undo();
	expect(controller.exportVsdx().bytes).toEqual(before);
	expect(controller.state.selectedShape?.id).toBe('1');
	await controller.redo();
	expect(controller.exportVsdx().bytes).toEqual(created);
	expect(controller.state.selectedShape?.id).toBe('2');
	controller.destroy();
});
it.each(['selection', 'page', 'zoom', 'layers', 'cancel', 'load', 'model'] as const)(
	'rejects earlier draft intent after %s even when view values return',
	async (kind) => {
		const controller = await setup();
		const token = controller.captureCreationToken('0');
		await change(controller, kind);
		expect(controller.isCreationTokenCurrent(token)).toBe(false);
		await expect(controller.applyCreationEdits([create], token)).rejects.toHaveProperty(
			'name',
			'AbortError',
		);
		if (kind !== 'model') expect(controller.state.document!.pages[0]!.shapes).toHaveLength(1);
		expect(controller.state.edit.canUndo).toBe(false);
		controller.destroy();
	},
);
it.each(['selection', 'page', 'zoom', 'layers', 'cancel', 'load', 'model'] as const)(
	'refuses pending creation after %s before source/history acceptance',
	async (kind) => {
		const deferredEdit = deferred();
		const controller = await setup(async (bytes, commands) => {
			await deferredEdit.waiting;
			return edit(bytes, commands);
		});
		const before = controller.exportVsdx().bytes;
		const token = controller.captureCreationToken('0');
		const pending = controller.applyCreationEdits([create], token);
		const rejected = expect(pending).rejects.toHaveProperty('name', 'AbortError');
		await change(controller, kind, before);
		deferredEdit.release();
		await rejected;
		if (kind !== 'model') expect(controller.exportVsdx().bytes).toEqual(before);
		expect(controller.state.document!.pages[0]!.shapes).toHaveLength(1);
		expect(controller.state.edit).toMatchObject({ canUndo: false, error: null });
		controller.destroy();
	},
);
it('rejects foreign/forged tokens, wrong-page and non-creation commands without source changes', async () => {
	const controller = await setup();
	const other = await setup();
	const before = controller.exportVsdx().bytes;
	await expect(
		controller.applyCreationEdits([create], other.captureCreationToken('0')),
	).rejects.toHaveProperty('name', 'AbortError');
	await expect(
		controller.applyCreationEdits([create], Object.freeze({}) as ViewerCreationToken),
	).rejects.toHaveProperty('name', 'AbortError');
	const token = controller.captureCreationToken('0');
	for (const commands of [
		[],
		[{ ...create, pageId: '1' }],
		[{ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 4 }] as VisioEdit[],
	])
		await expect(controller.applyCreationEdits(commands, token)).rejects.toThrow('Creation edits');
	expect(controller.exportVsdx().bytes).toEqual(before);
	controller.destroy();
	other.destroy();
});
it('rejects host option getters before editor invocation and hostile page getters during capture', async () => {
	const service = vi.fn(edit);
	const controller = await setup(service);
	const before = controller.exportVsdx().bytes;
	const token = controller.captureCreationToken('0');
	await expect(
		controller.applyCreationEdits(
			[
				{
					...create,
					get x() {
						controller.setZoom(2);
						return 4;
					},
				},
			],
			token,
		),
	).rejects.toHaveProperty('name', 'AbortError');
	expect(service).not.toHaveBeenCalled();
	expect(controller.exportVsdx().bytes).toEqual(before);
	const page = controller.state.document!.pages[0]!;
	Object.defineProperty(page, 'id', {
		configurable: true,
		get() {
			controller.clearSelection();
			return '0';
		},
	});
	expect(() => controller.captureCreationToken('0')).toThrow();
	controller.destroy();
});
it('preserves newer callback selection and rejects stale completion after atomic source acceptance', async () => {
	const controller = await setup();
	const token = controller.captureCreationToken('0');
	const primary = vi.fn();
	controller.onEvent((name, detail) => {
		if (name === 'shape-select') primary(detail);
	});
	let changed = false;
	controller.subscribe((state) => {
		if (!changed && state.document!.pages[0]!.shapes.length === 2 && !state.edit.busy) {
			changed = true;
			controller.clearSelection();
		}
	});
	await expect(controller.applyCreationEdits([create], token)).rejects.toHaveProperty(
		'name',
		'AbortError',
	);
	expect(controller.state.document!.pages[0]!.shapes).toHaveLength(2);
	expect(controller.state.selectedShapes).toEqual([]);
	expect(primary).toHaveBeenCalledOnce();
	expect(primary).toHaveBeenCalledWith(null);
	controller.destroy();
});
it('rejects disposed capture/current/completion and prevents late source acceptance', async () => {
	const deferredEdit = deferred();
	const controller = await setup(async (bytes, commands) => {
		await deferredEdit.waiting;
		return edit(bytes, commands);
	});
	const token = controller.captureCreationToken('0');
	const pending = controller.applyCreationEdits([create], token);
	const rejected = expect(pending).rejects.toHaveProperty('name', 'AbortError');
	controller.destroy();
	deferredEdit.release();
	await rejected;
	expect(controller.isCreationTokenCurrent(token)).toBe(false);
	expect(() => controller.captureCreationToken('0')).toThrow('destroyed');
});
