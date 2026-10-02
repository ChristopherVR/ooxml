import { expect, it, vi } from 'vitest';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import type { EditTransactionResult, CancellableEditor } from './worker-editor.js';
import { DocumentHistory } from './document-history.js';
const model = () => structuredClone(demoDocument);
const result = (n = 2): EditTransactionResult => ({
	bytes: new Uint8Array([n]),
	document: model(),
	changedParts: ['page.xml'],
	diagnostics: [{ code: 'warning', message: 'Unverified' }],
});
function deferred<T>() {
	let resolve!: (v: T) => void;
	let reject!: (e: Error) => void;
	const promise = new Promise<T>((a, b) => {
		resolve = a;
		reject = b;
	});
	return { promise, resolve, reject };
}
it.each(['cancel', 'load', 'replace', 'destroy'] as const)(
	'isolates a late undo parser completion after %s',
	async (action) => {
		const late = deferred<ReturnType<typeof model>>();
		const parser = Object.assign(
			vi
				.fn()
				.mockResolvedValueOnce(model())
				.mockImplementationOnce(() => late.promise)
				.mockResolvedValue(model()),
			{ cancel: vi.fn() },
		);
		const viewer = new ViewerController(
			parser,
			() => {},
			async () => result(),
		);
		await viewer.load(new Uint8Array([1]));
		await viewer.replacePlainText('1', 's1', 'new');
		const events = vi.fn();
		viewer.onEvent(events);
		const undo = viewer.undo();
		const rejection = expect(undo).rejects.toMatchObject({ name: 'AbortError' });
		if (action === 'cancel') viewer.cancelEdit();
		if (action === 'load') await viewer.load(new Uint8Array([9]));
		if (action === 'replace') viewer.setDocument(model());
		if (action === 'destroy') viewer.destroy();
		late.resolve(model());
		await rejection;
		expect(events.mock.calls.filter((call) => call[0] === 'document-change')).toHaveLength(0);
		if (action === 'cancel') {
			expect(viewer.exportVsdx().bytes[0]).toBe(2);
			expect(viewer.state.edit.canUndo).toBe(true);
			expect(viewer.state.edit.canRedo).toBe(false);
		}
		if (action === 'load') expect(viewer.exportVsdx().bytes[0]).toBe(9);
	},
);
it('undo parser mutation and rejection cannot corrupt accepted source/history', async () => {
	const parser = vi
		.fn()
		.mockResolvedValueOnce(model())
		.mockImplementationOnce((bytes: Uint8Array) => {
			bytes[0] = 99;
			throw new Error('bad reparse');
		})
		.mockResolvedValue(model());
	const viewer = new ViewerController(
		parser,
		() => {},
		async () => result(),
	);
	await viewer.load(new Uint8Array([1]));
	await viewer.replacePlainText('1', 's1', 'new');
	const accepted = viewer.state.document;
	await expect(viewer.undo()).rejects.toThrow('bad reparse');
	expect(viewer.state.document).toBe(accepted);
	expect(viewer.exportVsdx().bytes[0]).toBe(2);
	await viewer.undo();
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
	expect(viewer.state.edit.dirty).toBe(false);
});
it('cancel then restart rejects late completion without clearing the new busy state', async () => {
	const old = deferred<EditTransactionResult>(),
		next = deferred<EditTransactionResult>();
	const editor = vi
		.fn<CancellableEditor>()
		.mockImplementationOnce(() => old.promise)
		.mockImplementationOnce(() => next.promise);
	const viewer = new ViewerController(
		async () => model(),
		() => {},
		editor,
	);
	await viewer.load(new Uint8Array([1]));
	const a = viewer.replacePlainText('1', 's1', 'a');
	const rejection = expect(a).rejects.toMatchObject({ name: 'AbortError' });
	viewer.cancelEdit();
	const b = viewer.replacePlainText('1', 's1', 'b');
	old.resolve(result(3));
	await rejection;
	expect(viewer.state.edit.busy).toBe(true);
	next.resolve(result(4));
	await b;
	expect(viewer.exportVsdx().bytes[0]).toBe(4);
});
it('owns ArrayBuffer loads and protects source against parser and editor input mutations', async () => {
	const source = new Uint8Array([1]);
	const viewer = new ViewerController(
		async (bytes) => {
			new Uint8Array(bytes instanceof Uint8Array ? bytes.buffer : bytes)[0] = 9;
			return model();
		},
		() => {},
		async (bytes) => {
			bytes[0] = 8;
			return result();
		},
	);
	const load = viewer.load(source.buffer);
	source[0] = 7;
	await load;
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
	await viewer.replacePlainText('1', 's1', 'a');
	await viewer.undo();
	expect(viewer.exportVsdx().bytes[0]).toBe(1);
});
it('snapshots diagnostics and output bytes cannot mutate history', () => {
	const history = new DocumentHistory(new Uint8Array([1]));
	const bytes = new Uint8Array([2]);
	const notes = [{ code: 'first', message: 'first' }];
	history.append(bytes, notes);
	bytes[0] = 9;
	notes[0]!.code = 'changed';
	const exported = history.export();
	exported.bytes[0] = 7;
	expect(history.export().bytes[0]).toBe(2);
	expect(exported.diagnostics[0]!.code).toBe('first');
	expect(Object.isFrozen(exported.diagnostics)).toBe(true);
	expect(Object.isFrozen(exported.diagnostics[0])).toBe(true);
});
it('navigation during an edit keeps explicit target and the new navigation', async () => {
	const pending = deferred<EditTransactionResult>();
	const editor = vi.fn<CancellableEditor>(() => pending.promise);
	const viewer = new ViewerController(
		async () => model(),
		() => {},
		editor,
	);
	await viewer.load(new Uint8Array([1]));
	const edit = viewer.replacePlainText('background', 'same-id', 'a');
	viewer.setPage(1);
	viewer.setZoom(2);
	viewer.setSearchQuery('Start');
	pending.resolve(result());
	await edit;
	expect(editor.mock.calls[0]![1][0]).toMatchObject({ pageId: 'background', shapeId: 'same-id' });
	expect(viewer.state.pageIndex).toBe(1);
	expect(viewer.state.zoom).toBe(2);
	expect(viewer.state.search.query).toBe('Start');
});
