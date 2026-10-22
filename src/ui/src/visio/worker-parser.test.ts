import { describe, it, expect, vi, afterEach } from 'vitest';
import { createWorkerParser, type WorkerLike } from './worker-parser';
import { demoDocument } from 'ooxml-core/visio/ui';
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

it('times out a stalled converter worker and ignores late completion before a fresh import', async () => {
	vi.useFakeTimers();
	const first = fake(),
		second = fake();
	const factory = vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second);
	const parser = createWorkerParser(factory, 20);
	const stalled = parser(new ArrayBuffer(1));
	const late = first.onmessage;
	const rejected = expect(stalled).rejects.toThrow('isolated parsing limit');
	await vi.advanceTimersByTimeAsync(20);
	await rejected;
	const next = parser(new ArrayBuffer(1));
	late?.({ data: { ok: true, document: demoDocument } } as MessageEvent);
	expect(second.terminate).not.toHaveBeenCalled();
	second.onmessage?.({ data: { ok: true, document: demoDocument } } as MessageEvent);
	await expect(next).resolves.toBe(demoDocument);
	expect(first.terminate).toHaveBeenCalledOnce();
	expect(second.terminate).toHaveBeenCalledOnce();
});

describe('the worker kept between page reads', () => {
	const ok = (worker: WorkerLike) =>
		worker.onmessage?.({ data: { ok: true, document: demoDocument } } as MessageEvent);

	it('serves the next page read and is released by a whole read', async () => {
		const first = fake(),
			second = fake();
		const factory = vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second);
		const parser = createWorkerParser(factory);
		const one = parser.pages!(new Uint8Array([1]), ['5']);
		ok(first);
		await one;
		expect(first.terminate).not.toHaveBeenCalled();
		const two = parser.pages!(new Uint8Array([2]), ['5']);
		expect(factory).toHaveBeenCalledTimes(1);
		expect(first.postMessage).toHaveBeenCalledTimes(2);
		ok(first);
		await two;
		const whole = parser(new Uint8Array([3]));
		expect(first.terminate).toHaveBeenCalledOnce();
		expect(factory).toHaveBeenCalledTimes(2);
		ok(second);
		await whole;
		expect(second.terminate).toHaveBeenCalledOnce();
	});

	it('is ended by a failure and by dispose', async () => {
		const first = fake(),
			second = fake();
		const factory = vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second);
		const parser = createWorkerParser(factory);
		const failed = parser.pages!(new Uint8Array([1]), ['5']);
		first.onerror?.({ message: 'bad' } as ErrorEvent);
		await expect(failed).rejects.toThrow('bad');
		expect(first.terminate).toHaveBeenCalledOnce();
		const next = parser.pages!(new Uint8Array([1]), ['5']);
		expect(factory).toHaveBeenCalledTimes(2);
		ok(second);
		await next;
		parser.dispose?.();
		expect(second.terminate).toHaveBeenCalledOnce();
	});
});
