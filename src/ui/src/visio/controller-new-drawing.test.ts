import { expect, it, vi } from 'vitest';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { demoDocument } from 'ooxml-core/visio/ui';
import { ViewerController } from './controller';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import type { CancellableEditor } from './worker-editor';

const editor: CancellableEditor = async (bytes, commands) => {
	const result = await editVsdx(bytes, commands);
	return { ...result, document: await parseVsdx(result.bytes) };
};
it('creates a clean editable blank source, replacing old undo history and selection', async () => {
	const controller = new ViewerController(parseVsdx, () => {}, editor);
	await controller.load(await createVsdxFixture('Old'));
	controller.selectAll();
	controller.setSearchQuery('Old');
	await controller.replacePlainText('1', '1', 'Old changed');
	const loaded = vi.fn();
	controller.onEvent((name, value) => {
		if (name === 'document-load') loaded(value);
	});
	await controller.createBlankDrawing();
	const state = controller.state;
	expect(state.document?.format).toBe('vsdx');
	expect(state.document?.pages).toHaveLength(1);
	expect(state.document?.pages[0]).toMatchObject({
		id: '0',
		name: 'Page-1',
		width: 8.5,
		height: 11,
		shapes: [],
	});
	expect(state.selectedShapes).toEqual([]);
	expect(state.search.query).toBe('');
	expect(state.layerVisibilityOverrides).toEqual([]);
	expect(state.edit).toMatchObject({
		sourceAvailable: true,
		dirty: false,
		canUndo: false,
		canRedo: false,
		busy: false,
	});
	expect(loaded).toHaveBeenCalledOnce();
	expect(loaded).toHaveBeenLastCalledWith(state.document);
	const original = controller.exportVsdx();
	expect(original.dirty).toBe(false);
	expect((await parseVsdx(original.bytes)).pages[0]!.shapes).toEqual([]);
	await controller.applyEdits([
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 1, height: 1 },
	]);
	expect(controller.state.edit.dirty).toBe(true);
	expect((await parseVsdx(controller.exportVsdx().bytes)).pages[0]!.shapes).toHaveLength(1);
	await controller.undo();
	expect(controller.exportVsdx().bytes).toEqual(original.bytes);
	expect(controller.state.edit.dirty).toBe(false);
	controller.destroy();
});

it('uses physical custom dimensions and snapshots host options before waiting', async () => {
	const controller = new ViewerController(parseVsdx);
	const options = { width: 6, height: 4 };
	const pending = controller.createBlankDrawing(options);
	options.width = 9;
	await pending;
	expect(controller.state.document!.pages[0]).toMatchObject({
		width: 6,
		height: 4,
	});
	controller.destroy();
});

it('preserves current source on invalid factory options and reports the load failure', async () => {
	const controller = new ViewerController(parseVsdx);
	const bytes = await createVsdxFixture('Old');
	await controller.load(bytes);
	const before = controller.state.document;
	const error = vi.fn();
	controller.onEvent((name, value) => {
		if (name === 'document-error') error(value);
	});
	await expect(controller.createBlankDrawing({ width: NaN })).rejects.toThrow(
		'positive finite inches',
	);
	expect(controller.state.document).toBe(before);
	expect(controller.state.loading).toBe(false);
	expect(controller.exportVsdx().bytes).toEqual(Uint8Array.from(bytes));
	expect(error).toHaveBeenCalledOnce();
	controller.destroy();
});

it.each(['cancel', 'replace', 'destroy', 'newer-load'] as const)(
	'keeps %s from being overwritten by late blank parsing',
	async (action) => {
		let finish!: (document: Awaited<ReturnType<typeof parseVsdx>>) => void;
		const parser = vi.fn(parseVsdx);
		const controller = new ViewerController(parser);
		await controller.load(await createVsdxFixture('Old'));
		const old = controller.state.document;
		const loaded = vi.fn();
		controller.onEvent((name) => {
			if (name === 'document-load') loaded();
		});
		parser.mockImplementationOnce(
			() =>
				new Promise((done) => {
					finish = done;
				}),
		);
		const pending = controller.createBlankDrawing();
		await vi.waitFor(() => expect(parser).toHaveBeenCalledTimes(2));
		if (action === 'cancel') controller.cancelLoad();
		else if (action === 'replace') controller.setDocument(demoDocument);
		else if (action === 'destroy') controller.destroy();
		else await controller.load(await createVsdxFixture('Newest'));
		finish(demoDocument);
		await pending;
		if (action === 'cancel') expect(controller.state.document).toBe(old);
		else if (action === 'replace') expect(controller.state.document).toBe(demoDocument);
		else if (action === 'destroy') expect(controller.state.document).toBeNull();
		else expect(controller.state.document!.pages[0]!.shapes[0]!.text?.plainText).toBe('Newest');
		expect(loaded).toHaveBeenCalledTimes(action === 'newer-load' ? 1 : 0);
		controller.destroy();
	},
);

it('guards destroyed calls and host option getters that replace the document', async () => {
	const controller = new ViewerController(parseVsdx);
	await controller.createBlankDrawing({
		get width() {
			controller.setDocument(demoDocument);
			return 5;
		},
	});
	expect(controller.state.document).toBe(demoDocument);
	expect(controller.state.edit.sourceAvailable).toBe(false);
	controller.destroy();
	await expect(controller.createBlankDrawing()).rejects.toThrow('destroyed');
});

it('supersedes an in-flight old source edit without accepting it into the new drawing', async () => {
	let release!: () => void;
	const waiting = new Promise<void>((done) => {
		release = done;
	});
	const controller = new ViewerController(
		parseVsdx,
		() => {},
		async (bytes, commands) => {
			await waiting;
			return editor(bytes, commands);
		},
	);
	await controller.load(await createVsdxFixture('Old'));
	const pending = controller.replacePlainText('1', '1', 'Late edit');
	const superseded = expect(pending).rejects.toHaveProperty('name', 'AbortError');
	await controller.createBlankDrawing();
	const accepted = controller.state.document;
	const original = controller.exportVsdx().bytes;
	release();
	await superseded;
	expect(controller.state.document).toBe(accepted);
	expect(controller.exportVsdx().bytes).toEqual(original);
	expect(controller.state.edit).toMatchObject({ dirty: false, canUndo: false, error: null });
	controller.destroy();
});
