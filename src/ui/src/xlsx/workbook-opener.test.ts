import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, getCell, putCell } from 'ooxml-core/xlsx';
import { saveWorkbook } from 'ooxml-core/xlsx/load';
import {
	createWorkbookOpener,
	OpenTimeoutError,
	type OpenRequest,
	type OpenResponse,
	type WorkerLike,
} from './workbook-opener';

interface FakeWorker extends WorkerLike {
	sent: OpenRequest[];
	say(message: OpenResponse): void;
}

function fake(): FakeWorker {
	const worker: FakeWorker = {
		onmessage: null,
		onerror: null,
		sent: [],
		postMessage: vi.fn((message: OpenRequest) => void worker.sent.push(message)),
		terminate: vi.fn(),
		say: (message) => worker.onmessage?.({ data: message } as MessageEvent),
	};
	return worker;
}

/** A saved workbook whose B1 (`=A1*3`) holds a stale 99 and asks to be calculated on open. */
async function staleBytes(): Promise<Uint8Array> {
	const workbook = createWorkbook({ sheets: ['S'] });
	putCell(workbook.sheets[0]!, 0, 0, { value: 2 });
	putCell(workbook.sheets[0]!, 0, 1, { value: 99, formula: 'A1*3' });
	workbook.fullCalcOnLoad = true;
	return saveWorkbook(workbook, 'xlsx');
}

const limits = { readMs: 1000, calculateMs: 50 };

afterEach(() => vi.useRealTimers());

describe('workbook opener', () => {
	it('opens in a worker, sends a copy of the bytes and ends the worker', async () => {
		const worker = fake();
		const open = createWorkbookOpener(() => worker, limits);
		const bytes = new Uint8Array([1, 2, 3]);
		const opened = open(bytes, { fileName: 'a.xlsx' });
		worker.say({ kind: 'ready' });
		const request = worker.sent[0]!;
		expect(request).toMatchObject({ fileName: 'a.xlsx', calculate: true });
		expect(new Uint8Array(request.bytes)).toEqual(bytes);
		expect(request.bytes).not.toBe(bytes.buffer);
		const workbook = createWorkbook({ sheets: ['S'] });
		worker.say({ kind: 'read' });
		worker.say({ kind: 'done', workbook, calculated: true });
		await expect(opened).resolves.toEqual({
			workbook,
			calculated: true,
			calculationStopped: false,
		});
		expect(worker.terminate).toHaveBeenCalledOnce();
	});

	it('abandons a calculation past its limit and reopens without calculating', async () => {
		vi.useFakeTimers();
		const slow = fake();
		const again = fake();
		const factory = vi.fn().mockReturnValueOnce(slow).mockReturnValueOnce(again);
		const opened = createWorkbookOpener(factory, limits)(new Uint8Array([1]), {});
		slow.say({ kind: 'ready' });
		slow.say({ kind: 'read' });
		await vi.advanceTimersByTimeAsync(limits.calculateMs);
		expect(slow.terminate).toHaveBeenCalledOnce();
		expect(again.sent[0]).toMatchObject({ calculate: false });
		const workbook = createWorkbook({ sheets: ['S'] });
		again.say({ kind: 'ready' });
		again.say({ kind: 'read' });
		again.say({ kind: 'done', workbook, calculated: false });
		await expect(opened).resolves.toEqual({
			workbook,
			calculated: false,
			calculationStopped: true,
		});
	});

	it('fails an open whose reading runs past its limit', async () => {
		vi.useFakeTimers();
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(new Uint8Array([1]), {});
		const failed = expect(opened).rejects.toBeInstanceOf(OpenTimeoutError);
		worker.say({ kind: 'ready' });
		await vi.advanceTimersByTimeAsync(limits.readMs);
		await failed;
		expect(worker.terminate).toHaveBeenCalledOnce();
	});

	it("reports the loader's error with its code", async () => {
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(new Uint8Array([1]), {});
		worker.say({ kind: 'ready' });
		worker.say({ kind: 'error', message: 'A password is required.', code: 'password-required' });
		await expect(opened).rejects.toMatchObject({
			message: 'A password is required.',
			code: 'password-required',
		});
	});

	it('does not reopen on the page when the worker fails during an open', async () => {
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
		worker.say({ kind: 'ready' });
		worker.onerror?.({ message: 'out of memory' } as ErrorEvent);
		await expect(opened).rejects.toThrow('out of memory');
	});

	it('opens and calculates on the page when the worker cannot start', async () => {
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
		worker.onerror?.({ message: 'blocked' } as ErrorEvent);
		const { workbook, calculated, calculationStopped } = await opened;
		expect(calculated).toBe(true);
		expect(calculationStopped).toBe(false);
		expect(getCell(workbook.sheets[0]!, 0, 1)?.value).toBe(6);
	});

	it('opens on the page without workers', async () => {
		vi.stubGlobal('Worker', undefined);
		try {
			const { workbook, calculated } = await createWorkbookOpener()(await staleBytes(), {});
			expect(calculated).toBe(true);
			expect(getCell(workbook.sheets[0]!, 0, 1)?.value).toBe(6);
		} finally {
			vi.unstubAllGlobals();
		}
	});
});

describe('the workbook worker', () => {
	it('reads the file, runs the calculation on open and replies with the workbook', async () => {
		const replies: OpenResponse[] = [];
		const scope = {
			location: { origin: 'https://example.test' },
			postMessage: (message: OpenResponse) => void replies.push(message),
			onmessage: null as ((event: MessageEvent<OpenRequest>) => Promise<void>) | null,
		};
		vi.stubGlobal('self', scope);
		try {
			await import('./xlsx-open-worker');
			expect(replies).toEqual([{ kind: 'ready' }]);
			const bytes = await staleBytes();
			await scope.onmessage?.({
				data: { bytes: bytes.slice().buffer, calculate: true },
				origin: '',
			} as MessageEvent<OpenRequest>);
			expect(replies.map((reply) => reply.kind)).toEqual(['ready', 'read', 'done']);
			const done = replies[2] as Extract<OpenResponse, { kind: 'done' }>;
			expect(done.calculated).toBe(true);
			expect(getCell(structuredClone(done.workbook).sheets[0]!, 0, 1)?.value).toBe(6);
		} finally {
			vi.unstubAllGlobals();
		}
	});
});
