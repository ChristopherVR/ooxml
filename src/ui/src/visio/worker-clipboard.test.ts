import { afterEach, expect, it, vi } from 'vitest';
import { createWorkerClipboardCapture, type ClipboardWorkerLike } from './worker-clipboard';
import { VISIO_CLIPBOARD_MAX_CHARS } from 'ooxml-core/visio';

const fake = (): ClipboardWorkerLike => ({
	onmessage: null,
	onerror: null,
	postMessage: vi.fn(),
	terminate: vi.fn(),
});
afterEach(() => vi.useRealTimers());
it('copies borrowed bytes and IDs and releases the isolated worker after capture', async () => {
	const worker = fake(),
		bytes = new Uint8Array([1]),
		ids = ['2', '1'];
	const pending = createWorkerClipboardCapture(() => worker)(bytes, '1', ids);
	const [sent, transferred] = vi.mocked(worker.postMessage).mock.calls[0]!;
	bytes[0] = 8;
	ids[0] = '9';
	expect(new Uint8Array(sent.bytes)).toEqual(new Uint8Array([1]));
	expect(sent.shapeIds).toEqual(['2', '1']);
	expect(transferred).toEqual([sent.bytes]);
	worker.onmessage?.({ data: { ok: true, text: 'captured' } } as MessageEvent);
	await expect(pending).resolves.toBe('captured');
	expect(worker.terminate).toHaveBeenCalledOnce();
	expect(worker.onmessage).toBeNull();
	expect(worker.onerror).toBeNull();
});
it('cancels obsolete workers and ignores retained late callbacks', async () => {
	const first = fake(),
		second = fake();
	const capture = createWorkerClipboardCapture(
		vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second),
	);
	const pending = capture(new Uint8Array([1]), '1', ['1']);
	const rejected = expect(pending).rejects.toHaveProperty('name', 'AbortError');
	const late = first.onmessage;
	const next = capture(new Uint8Array([1]), '1', ['1']);
	await rejected;
	late?.({ data: { ok: true, text: 'stale' } } as MessageEvent);
	expect(second.terminate).not.toHaveBeenCalled();
	const cancelled = expect(next).rejects.toHaveProperty('name', 'AbortError');
	capture.cancel?.();
	await cancelled;
	expect(first.terminate).toHaveBeenCalledOnce();
	expect(second.terminate).toHaveBeenCalledOnce();
});
it('enforces worker timeout and cleanup', async () => {
	vi.useFakeTimers();
	const worker = fake();
	const pending = createWorkerClipboardCapture(() => worker, 20)(new Uint8Array([1]), '1', ['1']);
	const rejected = expect(pending).rejects.toThrow('time limit');
	await vi.advanceTimersByTimeAsync(20);
	await rejected;
	expect(worker.terminate).toHaveBeenCalledOnce();
});
it('contains factory, transfer and runtime failures', async () => {
	await expect(
		createWorkerClipboardCapture(() => {
			throw new Error('factory');
		})(new Uint8Array([1]), '1', ['1']),
	).rejects.toThrow('factory');
	const worker = fake();
	vi.mocked(worker.postMessage).mockImplementation(() => {
		throw new Error('transfer');
	});
	await expect(
		createWorkerClipboardCapture(() => worker)(new Uint8Array([1]), '1', ['1']),
	).rejects.toThrow('transfer');
	expect(worker.terminate).toHaveBeenCalledOnce();
	const other = fake();
	const pending = createWorkerClipboardCapture(() => other)(new Uint8Array([1]), '1', ['1']);
	other.onerror?.({ message: 'worker crash' } as ErrorEvent);
	await expect(pending).rejects.toThrow('worker crash');
	expect(other.terminate).toHaveBeenCalledOnce();
});
it('bounds input before construction and rejects invalid or oversized worker responses', async () => {
	const factory = vi.fn(fake),
		capture = createWorkerClipboardCapture(factory);
	await expect(capture(new Uint8Array(32 * 1024 * 1024 + 1), '1', ['1'])).rejects.toThrow(
		'Invalid',
	);
	await expect(capture(new Uint8Array([1]), '1', Array(1001).fill('1'))).rejects.toThrow('Invalid');
	await expect(capture(new Uint8Array([1]), '1', [''])).rejects.toThrow('Invalid');
	expect(factory).not.toHaveBeenCalled();
	for (const data of [
		{ ok: true, text: 9 },
		{ ok: true, text: 'x'.repeat(VISIO_CLIPBOARD_MAX_CHARS + 1) },
		{ ok: false, message: 'Unsupported', code: 'ClipboardUnsupportedShape' },
	]) {
		const pending = capture(new Uint8Array([1]), '1', ['1']);
		const worker = factory.mock.results.at(-1)!.value;
		worker.onmessage?.({ data } as MessageEvent);
		await expect(pending).rejects.toThrow();
		expect(worker.terminate).toHaveBeenCalledOnce();
	}
});
