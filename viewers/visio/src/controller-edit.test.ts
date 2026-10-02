import { expect, it, vi } from 'vitest';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import type { EditTransactionResult, CancellableEditor } from './worker-editor.js';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from '../tests/fixture.mjs';
function result(value = 2): EditTransactionResult {
	return {
		bytes: new Uint8Array([value]),
		document: structuredClone(demoDocument),
		changedParts: ['page.xml'],
		diagnostics: [{ code: 'edit-caches-not-recalculated', message: 'Unverified' }],
	};
}
function deferred<T>() {
	let resolve!: (value: T) => void;
	return {
		promise: new Promise<T>((r) => {
			resolve = r;
		}),
		resolve: (value: T) => resolve(value),
	};
}
function setup(editor: CancellableEditor = async () => result()) {
	return new ViewerController(
		async () => structuredClone(demoDocument),
		() => {},
		editor,
	);
}
it('keeps owned original source, atomic edits and defensive exports with undo/redo', async () => {
	const viewer = setup();
	const source = new Uint8Array([1]);
	const load = viewer.load(source);
	source[0] = 99;
	await load;
	expect(viewer.exportVsdx().bytes).toEqual(new Uint8Array([1]));
	const originalGeneration = viewer.documentGeneration;
	viewer.setZoom(2);
	viewer.setPage(1);
	await viewer.replacePlainText('0', '1', 'new');
	expect(viewer.state.edit).toMatchObject({
		dirty: true,
		canUndo: true,
		canRedo: false,
		busy: false,
	});
	expect(viewer.documentGeneration).toBe(originalGeneration + 1);
	expect(viewer.state.zoom).toBe(2);
	expect(viewer.state.pageIndex).toBe(1);
	const copy = viewer.exportVsdx();
	copy.bytes[0] = 99;
	expect(viewer.exportVsdx().bytes[0]).toBe(2);
	await viewer.undo();
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
	expect(viewer.state.edit).toMatchObject({ dirty: false, canRedo: true });
	await viewer.redo();
	expect(viewer.exportVsdx().bytes[0]).toBe(2);
	expect(viewer.state.edit.diagnostics[0]?.code).toBe('edit-caches-not-recalculated');
});
it('no-op preserves generation/history/redo; successful new edit clears redo', async () => {
	const editor = vi.fn<CancellableEditor>().mockResolvedValue(result());
	const viewer = setup(editor);
	await viewer.load(new Uint8Array([1]));
	await viewer.replacePlainText('0', '1', 'a');
	await viewer.undo();
	const generation = viewer.documentGeneration;
	editor.mockResolvedValueOnce({ ...result(), changedParts: [] });
	await viewer.replacePlainText('0', '1', 'original');
	expect(viewer.documentGeneration).toBe(generation);
	expect(viewer.state.edit.canRedo).toBe(true);
	await viewer.replacePlainText('0', '1', 'b');
	expect(viewer.state.edit.canRedo).toBe(false);
});
it('failed edit/reparse leaves original bytes/model/history intact and retains code', async () => {
	const error = Object.assign(new Error('Rich text cannot be edited'), {
		code: 'UNSUPPORTED_TEXT_EDIT',
	});
	const editor = vi
		.fn<CancellableEditor>()
		.mockRejectedValueOnce(error)
		.mockResolvedValueOnce({ ...result(), document: {} as never });
	const viewer = setup(editor);
	await viewer.load(new Uint8Array([1]));
	const document = viewer.state.document;
	await expect(viewer.replacePlainText('0', '1', 'a')).rejects.toBe(error);
	expect(viewer.state.edit.error).toBe(error);
	await expect(viewer.replacePlainText('0', '1', 'a')).rejects.toThrow();
	expect(viewer.state.document).toBe(document);
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
	expect(viewer.state.edit.canUndo).toBe(false);
});
it.each(['cancel', 'load', 'replace', 'destroy'] as const)(
	'invalidates in-flight edits on %s',
	async (action) => {
		const pending = deferred<EditTransactionResult>();
		const editor = Object.assign(
			vi.fn(() => pending.promise),
			{ cancel: vi.fn() },
		);
		const viewer = setup(editor);
		await viewer.load(new Uint8Array([1]));
		const events = vi.fn();
		viewer.onEvent(events);
		const edit = viewer.replacePlainText('0', '1', 'a');
		const rejected = expect(edit).rejects.toMatchObject({ name: 'AbortError' });
		await expect(viewer.replacePlainText('0', '1', 'b')).rejects.toThrow('in progress');
		expect(() => viewer.exportVsdx()).toThrow('Wait');
		if (action === 'cancel') viewer.cancelEdit();
		if (action === 'load') await viewer.load(new Uint8Array([3]));
		if (action === 'replace') viewer.setDocument(demoDocument);
		if (action === 'destroy') viewer.destroy();
		pending.resolve(result());
		await rejected;
		expect(events.mock.calls.some((call) => call[0] === 'document-change')).toBe(false);
		if (action === 'cancel') expect(viewer.exportVsdx().bytes[0]).toBe(1);
		if (action === 'load') expect(viewer.exportVsdx().bytes[0]).toBe(3);
	},
);
it('reentrant cancellation before the worker starts prevents the transaction', async () => {
	const editor = vi.fn<CancellableEditor>().mockResolvedValue(result());
	const viewer = setup(editor);
	await viewer.load(new Uint8Array([1]));
	viewer.subscribe((state) => {
		if (state.edit.busy) viewer.cancelEdit();
	});
	await expect(viewer.replacePlainText('0', '1', 'a')).rejects.toMatchObject({
		name: 'AbortError',
	});
	expect(editor).not.toHaveBeenCalled();
});
it('reentrant source replacement on edit commit suppresses obsolete event', async () => {
	const viewer = setup();
	await viewer.load(new Uint8Array([1]));
	const events = vi.fn();
	viewer.onEvent(events);
	viewer.subscribe((state) => {
		if (state.edit.dirty) viewer.setDocument(demoDocument);
	});
	await viewer.replacePlainText('0', '1', 'a');
	expect(viewer.state.edit.sourceAvailable).toBe(false);
	expect(events).not.toHaveBeenCalled();
});
it('model-only and destroyed controllers cannot edit/save', async () => {
	const viewer = setup();
	viewer.setDocument(demoDocument);
	await expect(viewer.undo()).rejects.toThrow('Load');
	expect(() => viewer.exportVsdx()).toThrow('Load');
	viewer.destroy();
	await expect(viewer.redo()).rejects.toThrow('destroyed');
	expect(() => viewer.exportVsdx()).toThrow('destroyed');
});
it('failed replacement load retains an editable previously accepted source', async () => {
	const parser = vi
		.fn()
		.mockResolvedValueOnce(demoDocument)
		.mockRejectedValueOnce(new Error('Bad ZIP'));
	const viewer = new ViewerController(
		parser,
		() => {},
		async () => result(),
	);
	await viewer.load(new Uint8Array([1]));
	await expect(viewer.load(new Uint8Array([2]))).rejects.toThrow('Bad ZIP');
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
	await viewer.replacePlainText('0', '1', 'a');
	expect(viewer.state.edit.dirty).toBe(true);
});
it('real core transaction edits, reparses, undoes and redoes source-backed bytes', async () => {
	const editor: CancellableEditor = async (bytes, commands) => {
		const edited = await editVsdx(bytes, commands);
		return { ...edited, document: await parseVsdx(edited.bytes) };
	};
	const viewer = new ViewerController(parseVsdx, () => {}, editor);
	const bytes = await createVsdxFixture('Original');
	await viewer.load(bytes);
	const page = viewer.state.document!.pages[0]!;
	const shape = page.shapes[0]!;
	await viewer.replacePlainText(page.id, shape.id, '<script>inert & text</script>');
	expect(viewer.state.document!.pages[0]!.shapes[0]!.text.plainText).toBe(
		'<script>inert & text</script>',
	);
	expect((await parseVsdx(viewer.exportVsdx().bytes)).pages[0]!.shapes[0]!.text.plainText).toBe(
		'<script>inert & text</script>',
	);
	await viewer.undo();
	expect(viewer.exportVsdx().bytes).toEqual(Uint8Array.from(bytes));
	await viewer.redo();
	expect(viewer.state.edit.dirty).toBe(true);
});
