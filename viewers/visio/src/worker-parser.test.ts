import { describe, it, expect, vi, afterEach } from 'vitest';
import { createWorkerParser, type WorkerLike } from './worker-parser.js';
import { demoDocument } from './demo-document.js';
function fake(): WorkerLike {
	return { onmessage: null, onerror: null, postMessage: vi.fn(), terminate: vi.fn() };
}
afterEach(() => vi.useRealTimers());
describe('isolated parser lifecycle', () => {
	it('transfers a copy and releases the worker on success', async () => {
		const worker = fake(),
			parser = createWorkerParser(() => worker);
		const bytes = new Uint8Array([1, 2]);
		const load = parser(bytes);
		expect(worker.postMessage).toHaveBeenCalled();
		const sent = vi.mocked(worker.postMessage).mock.calls[0]![0];
		expect(sent).not.toBe(bytes.buffer);
		worker.onmessage?.({ data: { ok: true, document: demoDocument } } as MessageEvent);
		await expect(load).resolves.toBe(demoDocument);
		expect(worker.terminate).toHaveBeenCalledTimes(1);
		expect(worker.onmessage).toBeNull();
	});
	it('rejects and terminates a superseded parse', async () => {
		const worker = fake(),
			parser = createWorkerParser(() => worker);
		const load = parser(new ArrayBuffer(0));
		parser.cancel?.();
		await expect(load).rejects.toMatchObject({ name: 'AbortError' });
		expect(worker.terminate).toHaveBeenCalledTimes(1);
	});
	it('enforces a parent-side hard timeout', async () => {
		vi.useFakeTimers();
		const worker = fake(),
			parser = createWorkerParser(() => worker, 20);
		const load = parser(new ArrayBuffer(0));
		const rejection = expect(load).rejects.toThrow('isolated parsing limit');
		await vi.advanceTimersByTimeAsync(20);
		await rejection;
		expect(worker.terminate).toHaveBeenCalledOnce();
	});
	it('reports worker errors without unsafe fallback parsing', async () => {
		const worker = fake(),
			parser = createWorkerParser(() => worker);
		const load = parser(new ArrayBuffer(0));
		worker.onerror?.({ message: 'Worker blocked' } as ErrorEvent);
		await expect(load).rejects.toThrow('Worker blocked');
		expect(worker.terminate).toHaveBeenCalledOnce();
	});
});
