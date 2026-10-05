import { afterEach, expect, it, vi } from 'vitest';
import { createWorkerEditor, type EditWorkerLike } from './worker-editor.js';
import { demoDocument } from './demo-document.js';
const command = { type: 'replace-plain-text' as const, pageId: '1', shapeId: '1', text: 'new' };
const fake = (): EditWorkerLike => ({
	onmessage: null,
	onerror: null,
	postMessage: vi.fn(),
	terminate: vi.fn(),
});
const valid = () => ({
	ok: true,
	bytes: new Uint8Array([2]),
	document: demoDocument,
	changedParts: ['visio/pages/page1.xml'],
	diagnostics: [],
});
afterEach(() => vi.useRealTimers());
it('independent review: transfers only its exact owned source copy', async () => {
	const worker = fake();
	const source = new Uint8Array([99, 1, 2, 99]);
	const pending = createWorkerEditor(() => worker)(source.subarray(1, 3), [command]);
	const [message, transfers] = vi.mocked(worker.postMessage).mock.calls[0]!;
	expect(new Uint8Array(message.bytes)).toEqual(new Uint8Array([1, 2]));
	expect(transfers).toEqual([message.bytes]);
	expect(message.bytes).not.toBe(source.buffer);
	structuredClone(message, { transfer: transfers });
	expect(source).toEqual(new Uint8Array([99, 1, 2, 99]));
	worker.onmessage?.({ data: valid() } as MessageEvent);
	await pending;
});
it('independent review: rejects excessive command payload before worker allocation', async () => {
	for (const commands of [
		Array.from({ length: 1001 }, () => command),
		[{ ...command, text: 'x'.repeat(1_000_001) }],
		[{ ...command, shapeId: 'x'.repeat(257) }],
	]) {
		const factory = vi.fn(fake);
		await expect(
			Promise.resolve().then(() => createWorkerEditor(factory)(new Uint8Array([1]), commands)),
		).rejects.toThrow();
		expect(factory).not.toHaveBeenCalled();
	}
});
it('independent review: ignores extra command members rather than invoking their getters', async () => {
	const worker = fake();
	const malicious = {
		...command,
		get unused() {
			throw new Error('Must not read unrelated members');
		},
	};
	const pending = createWorkerEditor(() => worker)(new Uint8Array([1]), [malicious]);
	expect(vi.mocked(worker.postMessage).mock.calls[0]![0].edits).toEqual([command]);
	worker.onmessage?.({ data: valid() } as MessageEvent);
	await pending;
});
it('independent review: response accessors fail closed immediately', async () => {
	const worker = fake();
	const pending = createWorkerEditor(() => worker)(new Uint8Array([1]), [command]);
	const rejection = expect(pending).rejects.toThrow();
	expect(() =>
		worker.onmessage?.({
			data: new Proxy(
				{},
				{
					get() {
						throw new Error('Bad response accessor');
					},
				},
			),
		} as MessageEvent),
	).not.toThrow();
	await rejection;
	expect(worker.terminate).toHaveBeenCalledOnce();
});
it('independent review: terminal cleanup exceptions cannot orphan the promise', async () => {
	const worker = fake();
	vi.mocked(worker.terminate).mockImplementation(() => {
		throw new Error('Termination unavailable');
	});
	const pending = createWorkerEditor(() => worker)(new Uint8Array([1]), [command]);
	expect(() => worker.onmessage?.({ data: valid() } as MessageEvent)).not.toThrow();
	await expect(pending).rejects.toThrow('Termination unavailable');
});
it('independent review: unsupported browser cannot fall back to main-thread editing', async () => {
	vi.stubGlobal('Worker', undefined);
	try {
		await expect(createWorkerEditor()(new Uint8Array([1]), [command])).rejects.toThrow(
			'worker support',
		);
	} finally {
		vi.unstubAllGlobals();
	}
});
