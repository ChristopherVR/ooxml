import { afterEach, expect, it, vi } from 'vitest';
import { createWorkerEditor, type EditWorkerLike } from './worker-editor.js';
import { demoDocument } from './demo-document.js';
function fake(): EditWorkerLike {
	return { onmessage: null, onerror: null, postMessage: vi.fn(), terminate: vi.fn() };
}
const command = { type: 'replace-plain-text' as const, pageId: '0', shapeId: '1', text: 'text' };
const response = () => ({
	ok: true,
	bytes: new Uint8Array([2]),
	document: demoDocument,
	changedParts: ['page.xml'],
	diagnostics: [],
});
afterEach(() => vi.useRealTimers());
it('copies borrowed source/commands and cleans up success', async () => {
	const worker = fake(),
		edit = createWorkerEditor(() => worker),
		source = new Uint8Array([1]),
		commands = [{ ...command }];
	const pending = edit(source, commands);
	const sent = vi.mocked(worker.postMessage).mock.calls[0]![0];
	source[0] = 8;
	commands[0]!.text = 'changed';
	expect(new Uint8Array(sent.bytes)[0]).toBe(1);
	expect(sent.edits[0]).toMatchObject({ text: 'text' });
	worker.onmessage?.({ data: response() } as MessageEvent);
	await expect(pending).resolves.toMatchObject({ bytes: new Uint8Array([2]) });
	expect(worker.terminate).toHaveBeenCalledOnce();
	expect(worker.onmessage).toBeNull();
});
it('cancels and ignores late responses from an obsolete worker', async () => {
	const first = fake(),
		second = fake(),
		edit = createWorkerEditor(vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second));
	const pending = edit(new Uint8Array([1]), [command]);
	const rejected = expect(pending).rejects.toMatchObject({ name: 'AbortError' });
	const late = first.onmessage;
	const next = edit(new Uint8Array([1]), [command]);
	await rejected;
	late?.({ data: response() } as MessageEvent);
	expect(second.terminate).not.toHaveBeenCalled();
	second.onmessage?.({ data: response() } as MessageEvent);
	await next;
	expect(first.terminate).toHaveBeenCalledOnce();
});
it('terminates workers at the hard timeout', async () => {
	vi.useFakeTimers();
	const worker = fake(),
		edit = createWorkerEditor(() => worker, 20);
	const pending = edit(new Uint8Array([1]), [command]);
	const rejected = expect(pending).rejects.toThrow('time limit');
	await vi.advanceTimersByTimeAsync(20);
	await rejected;
	expect(worker.terminate).toHaveBeenCalledOnce();
});
it('preserves core error code and rejects malformed result without fallback', async () => {
	for (const data of [
		{ ok: false, code: 'UNSUPPORTED_TEXT_EDIT', message: 'Fields unsupported' },
		{ ok: true, bytes: 'bad' },
		{ ...response(), diagnostics: [{}] },
		{ ...response(), changedParts: [1] },
	]) {
		const worker = fake(),
			edit = createWorkerEditor(() => worker);
		const pending = edit(new Uint8Array([1]), [command]);
		worker.onmessage?.({ data } as MessageEvent);
		if ('code' in data)
			await expect(pending).rejects.toMatchObject({ code: 'UNSUPPORTED_TEXT_EDIT' });
		else await expect(pending).rejects.toThrow('invalid response');
		expect(worker.terminate).toHaveBeenCalledOnce();
	}
});
it('reports constructor, postMessage and runtime errors without fallback', async () => {
	await expect(
		createWorkerEditor(() => {
			throw new Error('CSP blocked');
		})(new Uint8Array([1]), [command]),
	).rejects.toThrow('CSP blocked');
	const worker = fake();
	vi.mocked(worker.postMessage).mockImplementation(() => {
		throw new Error('transfer failed');
	});
	await expect(createWorkerEditor(() => worker)(new Uint8Array([1]), [command])).rejects.toThrow(
		'transfer failed',
	);
	expect(worker.terminate).toHaveBeenCalledOnce();
	const other = fake(),
		pending = createWorkerEditor(() => other)(new Uint8Array([1]), [command]);
	other.onerror?.({ message: 'Crashed' } as ErrorEvent);
	await expect(pending).rejects.toThrow('Crashed');
});
it('enforces source and worker-output ceilings', async () => {
	const factory = vi.fn(fake);
	await expect(
		createWorkerEditor(factory)(new Uint8Array(32 * 1024 * 1024 + 1), [command]),
	).rejects.toThrow('32 MiB');
	expect(factory).not.toHaveBeenCalled();
	const worker = fake(),
		pending = createWorkerEditor(() => worker)(new Uint8Array([1]), [command]);
	worker.onmessage?.({
		data: { ...response(), bytes: new Uint8Array(32 * 1024 * 1024 + 1) },
	} as MessageEvent);
	await expect(pending).rejects.toThrow('invalid response');
});
it('rejects unbounded commands before worker construction and strips unknown properties', async () => {
	const factory = vi.fn(fake),
		editor = createWorkerEditor(factory);
	await expect(
		editor(new Uint8Array([1]), [{ ...command, text: 'x'.repeat(1_000_001) }]),
	).rejects.toThrow('limit');
	await expect(
		editor(
			new Uint8Array([1]),
			Array.from({ length: 1001 }, () => command),
		),
	).rejects.toThrow('1000');
	await expect(
		editor(new Uint8Array([1]), [{ ...command, pageId: 'x'.repeat(257) }]),
	).rejects.toThrow('Invalid');
	expect(factory).not.toHaveBeenCalled();
	const pending = editor(new Uint8Array([1]), [{ ...command, extra: () => {} } as typeof command]);
	const worker = factory.mock.results[0]!.value;
	expect(vi.mocked(worker.postMessage).mock.calls[0]![0].edits[0]).not.toHaveProperty('extra');
	worker.onmessage?.({ data: response() } as MessageEvent);
	await pending;
});

it('bounds and snapshots geometry commands without cloning arbitrary extras', async () => {
	const worker = fake(),
		factory = vi.fn(() => worker),
		editor = createWorkerEditor(factory);
	await expect(
		editor(new Uint8Array([1]), [{ type: 'move-shape', pageId: '1', shapeId: '2', x: NaN, y: 1 }]),
	).rejects.toThrow('Invalid');
	expect(factory).not.toHaveBeenCalled();
	const move = {
		type: 'move-shape' as const,
		pageId: '1',
		shapeId: '2',
		x: -2,
		y: 3,
		extra: () => {},
	};
	const pending = editor(new Uint8Array([1]), [
		move,
		{
			type: 'create-rectangle',
			pageId: '1',
			shapeId: '3',
			x: 1,
			y: 2,
			width: 3,
			height: 4,
			text: 'Rectangle',
		},
	]);
	move.x = 99;
	const sent = vi.mocked(worker.postMessage).mock.calls[0]![0];
	expect(sent.edits[0]).toEqual({ type: 'move-shape', pageId: '1', shapeId: '2', x: -2, y: 3 });
	expect(sent.edits[1]).toMatchObject({ type: 'create-rectangle', text: 'Rectangle' });
	worker.onmessage?.({ data: response() } as MessageEvent);
	await pending;
});
