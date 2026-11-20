import { afterEach, describe, expect, it, vi } from 'vitest';
import { calculateOnOpenCells, createWorkbook, getCell, putCell } from 'ooxml-core/xlsx';
import { loadWorkbook, saveWorkbook } from 'ooxml-core/xlsx/load';
import {
	createWorkbookOpener,
	type CalculateRequest,
	type CalculateResponse,
	type WorkerLike,
} from './workbook-opener';

interface FakeWorker extends WorkerLike {
	sent: CalculateRequest[];
	say(message: CalculateResponse): void;
}

function fake(): FakeWorker {
	const worker: FakeWorker = {
		onmessage: null,
		onerror: null,
		sent: [],
		postMessage: vi.fn((message: CalculateRequest) => void worker.sent.push(message)),
		terminate: vi.fn(),
		say: (message) => worker.onmessage?.({ data: message } as MessageEvent),
	};
	return worker;
}

/**
 * A saved workbook whose B1 (6) holds a stale 99, asking to be calculated on open or not. Its
 * matrix product needs a little work, more than the page is given in these tests.
 */
async function staleBytes(calculate = true): Promise<Uint8Array> {
	const workbook = createWorkbook({ sheets: ['S'] });
	putCell(workbook.sheets[0]!, 0, 0, { value: 2 });
	putCell(workbook.sheets[0]!, 0, 1, { value: 99, formula: 'A1*3+SUM(MMULT({1},{0}))' });
	workbook.fullCalcOnLoad = calculate;
	return saveWorkbook(workbook, 'xlsx');
}

/** What the worker would post back for `bytes`. */
async function calculated(bytes: Uint8Array): Promise<CalculateResponse> {
	const { cells } = calculateOnOpenCells(await loadWorkbook(bytes.slice()));
	return { kind: 'done', cells: cells! };
}

const limits = { pageMs: 250, pageWorkCap: 0, startMs: 500, readMs: 1000, calculateMs: 50 };
const b1 = (workbook: { sheets: { rows: Map<number, Map<number, unknown>> }[] }) =>
	getCell(workbook.sheets[0] as never, 0, 1)?.value;

/** Waits until the opener has sent `worker` its `count`th request. */
async function requested(worker: FakeWorker, count = 1): Promise<CalculateRequest> {
	await vi.waitFor(() => expect(worker.sent.length).toBeGreaterThanOrEqual(count));
	return worker.sent[count - 1]!;
}

/** Fakes only the opener's own timers, so the zip reader still yields as usual. */
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });

afterEach(() => vi.useRealTimers());

describe('workbook opener', () => {
	it('reads a file that asks for no calculation on the page, without a worker', async () => {
		const factory = vi.fn(fake);
		const { workbook, calculated, calculationStopped } = await createWorkbookOpener(
			factory,
			limits,
		)(await staleBytes(false), {});
		expect(factory).not.toHaveBeenCalled();
		expect([b1(workbook), calculated, calculationStopped]).toEqual([99, false, false]);
	});

	it('calculates a quick workbook on the page, without a worker', async () => {
		const factory = vi.fn(fake);
		const { workbook, calculated, calculationStopped } = await createWorkbookOpener(factory, {
			...limits,
			pageWorkCap: 1_000,
		})(await staleBytes(), {});
		expect(factory).not.toHaveBeenCalled();
		expect([b1(workbook), calculated, calculationStopped]).toEqual([6, true, false]);
	});

	it('leaves a slow calculation to a worker sent a copy of the bytes', async () => {
		const worker = fake();
		const bytes = await staleBytes();
		const opened = createWorkbookOpener(() => worker, limits)(bytes, { fileName: 'a.xlsx' });
		const request = await requested(worker);
		expect(request).toMatchObject({ fileName: 'a.xlsx', timeLimitMs: limits.calculateMs });
		expect(new Uint8Array(request.bytes)).toEqual(bytes);
		expect(request.bytes).not.toBe(bytes.buffer);
		worker.say({ kind: 'ready' });
		worker.say({ kind: 'read' });
		worker.say(await calculated(bytes));
		const { workbook, calculated: ran, calculationStopped } = await opened;
		expect([b1(workbook), ran, calculationStopped]).toEqual([6, true, false]);
		expect(worker.terminate).not.toHaveBeenCalled();
	});

	it('keeps the worker for the next calculation and ends it once idle', async () => {
		fakeTimers();
		const worker = fake();
		const factory = vi.fn(() => worker);
		const open = createWorkbookOpener(factory, limits);
		const bytes = await staleBytes();
		const done = await calculated(bytes);
		for (const [i, name] of ['a.xlsx', 'b.xlsx'].entries()) {
			const opened = open(bytes, { fileName: name });
			await requested(worker, i + 1);
			worker.say({ kind: 'read' });
			worker.say(done);
			await opened;
		}
		expect(factory).toHaveBeenCalledOnce();
		expect(worker.sent.map((request) => request.fileName)).toEqual(['a.xlsx', 'b.xlsx']);
		expect(worker.terminate).not.toHaveBeenCalled();
		await vi.advanceTimersByTimeAsync(30_000);
		expect(worker.terminate).toHaveBeenCalledOnce();
	});

	it('shows the saved values when the calculation stops at its limit', async () => {
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
		await requested(worker);
		worker.say({ kind: 'read' });
		worker.say({ kind: 'stopped', reason: 'time' });
		const { workbook, calculated, calculationStopped } = await opened;
		expect([b1(workbook), calculated, calculationStopped]).toEqual([99, false, true]);
		expect(worker.terminate).not.toHaveBeenCalled();
	});

	it('ends a worker that does not stop in time and shows the saved values', async () => {
		fakeTimers();
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
		await requested(worker);
		worker.say({ kind: 'read' });
		await vi.advanceTimersByTimeAsync(limits.calculateMs + 5_000);
		const { workbook, calculationStopped } = await opened;
		expect([b1(workbook), calculationStopped]).toEqual([99, true]);
		expect(worker.terminate).toHaveBeenCalledOnce();
	});

	it('calculates on the page when the worker fails or cannot start', async () => {
		for (const fail of [
			(worker: FakeWorker) => worker.say({ kind: 'error', message: 'out of memory' }),
			(worker: FakeWorker) => worker.onerror?.({ message: 'blocked' } as ErrorEvent),
		]) {
			const worker = fake();
			const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
			await requested(worker);
			fail(worker);
			const { workbook, calculated, calculationStopped } = await opened;
			expect([b1(workbook), calculated, calculationStopped]).toEqual([6, true, false]);
			expect(worker.terminate).toHaveBeenCalledOnce();
		}
	});

	it('calculates on the page when a new worker never starts', async () => {
		fakeTimers();
		const worker = fake();
		const opened = createWorkbookOpener(() => worker, limits)(await staleBytes(), {});
		await requested(worker);
		await vi.advanceTimersByTimeAsync(limits.startMs);
		const { workbook, calculated } = await opened;
		expect([b1(workbook), calculated]).toEqual([6, true]);
		expect(worker.terminate).toHaveBeenCalledOnce();
	});

	it('calculates on the page without workers', async () => {
		vi.stubGlobal('Worker', undefined);
		try {
			const open = createWorkbookOpener(undefined, limits);
			const { workbook, calculated } = await open(await staleBytes(), {});
			expect([b1(workbook), calculated]).toEqual([6, true]);
		} finally {
			vi.unstubAllGlobals();
		}
	});
});

describe('the workbook worker', () => {
	it('reads the file, calculates it and replies with the cells it wrote', async () => {
		const replies: CalculateResponse[] = [];
		const transfers: Transferable[][] = [];
		const scope = {
			location: { origin: 'https://example.test' },
			postMessage: (message: CalculateResponse, transfer: Transferable[]) => {
				replies.push(message);
				transfers.push(transfer);
			},
			onmessage: null as ((event: MessageEvent<CalculateRequest>) => Promise<void>) | null,
		};
		vi.stubGlobal('self', scope);
		try {
			await import('./xlsx-open-worker');
			expect(replies).toEqual([{ kind: 'ready' }]);
			const bytes = await staleBytes();
			await scope.onmessage?.({
				data: { bytes: bytes.slice().buffer, timeLimitMs: 1000 },
				origin: '',
			} as MessageEvent<CalculateRequest>);
			expect(replies.map((reply) => reply.kind)).toEqual(['ready', 'read', 'done']);
			const done = replies[2] as Extract<CalculateResponse, { kind: 'done' }>;
			expect(done).toEqual(await calculated(bytes));
			expect(transfers[2]).toContain(done.cells.sheets[0]!.cells.buffer);
		} finally {
			vi.unstubAllGlobals();
		}
	});
});
