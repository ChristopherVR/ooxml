/**
 * Opens workbook files. The page reads every file itself, as it always has. A file that asks for a
 * full calculation on open (`calcPr fullCalcOnLoad`) is first calculated on the page, briefly: for
 * at most as long as reading it took (and no less than a quarter of a second), with no formula
 * allowed more than a quarter of a second of work. Most workbooks finish there, as fast as before.
 * One that does not is sent to a worker (`xlsx-open-worker.ts`), which reads the file again,
 * runs the whole calculation off the page under a time limit and posts back only the cells it
 * wrote, so slow formulas never freeze the editor for long and nothing large is copied between
 * threads. Past the limit the formulas it did not finish show the values the file was saved
 * with. Where workers are unavailable (no `Worker`, or the worker script cannot start) the
 * calculation runs on the page, under the same time limit.
 */
import {
	applyCalculatedCells,
	calculateOnOpen,
	needsCalculationOnOpen,
	type CalculatedCells,
	type CalculationStop,
	type OpenCalculationLimits,
	type Workbook,
} from 'ooxml-core/xlsx';

/** What the worker is asked to calculate. */
export interface CalculateRequest {
	bytes: ArrayBuffer;
	fileName?: string;
	password?: string;
	/** How long the calculation may run before it stops. */
	timeLimitMs: number;
}

/** What the worker tells: it started, it read the file, or how the calculation ended. */
export type CalculateResponse =
	| { kind: 'ready' }
	| { kind: 'read' }
	| { kind: 'done'; cells: CalculatedCells }
	/** The file does not ask for a calculation after all. */
	| { kind: 'skipped' }
	| { kind: 'stopped'; reason: CalculationStop }
	| { kind: 'error'; message: string };

export interface OpenOptions {
	fileName?: string;
	password?: string;
}

export interface OpenedWorkbook {
	workbook: Workbook;
	/** Whether the calculation the file asks for on open ran to the end. */
	calculated: boolean;
	/** It stopped early, so some or all formulas show the values the file was saved with. */
	calculationStopped: boolean;
}

export type WorkbookOpener = (
	bytes: Uint8Array | ArrayBuffer,
	options: OpenOptions,
) => Promise<OpenedWorkbook>;

export interface WorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: CalculateRequest, transfer: Transferable[]): void;
	terminate(): void;
}
export type WorkerFactory = () => WorkerLike;

export interface OpenLimits {
	/** The least time the page may calculate before it leaves the rest to a worker. */
	pageMs: number;
	/** The most work one formula may need on the page (units of `MAX_FORMULA_WORK`). */
	pageWorkCap: number;
	/** How long a new worker may take to start; past it the page calculates instead. */
	startMs: number;
	/** How long the worker may take to read the file before its calculation is abandoned. */
	readMs: number;
	/** How long the calculation on open may run. */
	calculateMs: number;
}

export const DEFAULT_OPEN_LIMITS: OpenLimits = {
	pageMs: 250,
	pageWorkCap: 250_000_000,
	startMs: 10_000,
	readMs: 60_000,
	calculateMs: 10_000,
};

/**
 * How long past its time limit a calculation may go before the worker is ended: the worker stops
 * itself between formulas, and each formula is bounded by its own work budget.
 */
const STOP_GRACE_MS = 5_000;

type Outcome =
	| { kind: 'done'; cells: CalculatedCells }
	| { kind: 'skipped' }
	| { kind: 'stopped' }
	/** No worker could run it: the page calculates instead. */
	| { kind: 'unavailable' };

const loadOptions = (options: OpenOptions) => ({
	...(options.fileName === undefined ? {} : { fileName: options.fileName }),
	...(options.password === undefined ? {} : { password: options.password }),
});

/** How long a worker that finished a calculation stays for the next one. */
const IDLE_MS = 30_000;

/**
 * The worker kept between opens: starting one loads the whole reader, which would otherwise be
 * paid again by every file calculated in a row. A failed or abandoned run ends its worker.
 */
class WorkerPool {
	private idle: WorkerLike | undefined;
	private idleTimer: ReturnType<typeof setTimeout> | undefined;
	/** Workers whose script has run (they said `ready`). */
	private readonly started = new WeakSet<WorkerLike>();

	constructor(private readonly create: WorkerFactory) {}

	/** An idle worker or a new one; undefined when none can start. */
	acquire(): WorkerLike | undefined {
		clearTimeout(this.idleTimer);
		const idle = this.idle;
		this.idle = undefined;
		if (idle) return idle;
		try {
			return this.create();
		} catch {
			return undefined;
		}
	}

	/** Whether `worker` has started; one still starting says so while it waits. */
	hasStarted(worker: WorkerLike): boolean {
		return this.started.has(worker);
	}

	markStarted(worker: WorkerLike): void {
		this.started.add(worker);
	}

	release(worker: WorkerLike): void {
		worker.onmessage = (event) => {
			if ((event.data as CalculateResponse).kind === 'ready') this.started.add(worker);
		};
		worker.onerror = null;
		if (this.idle) return worker.terminate();
		this.idle = worker;
		this.idleTimer = setTimeout(() => {
			if (this.idle !== worker) return;
			this.idle = undefined;
			worker.terminate();
		}, IDLE_MS);
	}

	discard(worker: WorkerLike): void {
		worker.onmessage = null;
		worker.onerror = null;
		worker.terminate();
	}
}

/** Sends a copy of the bytes to a worker that reads them and runs the calculation on open. */
function calculateInWorker(
	pool: WorkerPool,
	bytes: Uint8Array | ArrayBuffer,
	options: OpenOptions,
	limits: OpenLimits,
): Promise<Outcome> {
	return new Promise((resolve) => {
		const worker = pool.acquire();
		if (!worker) return resolve({ kind: 'unavailable' });
		let timer: ReturnType<typeof setTimeout> | undefined;
		const finish = (outcome: Outcome, discard = false) => {
			clearTimeout(timer);
			if (discard) pool.discard(worker);
			else pool.release(worker);
			resolve(outcome);
		};
		const arm = (ms: number) => {
			clearTimeout(timer);
			timer = setTimeout(() => finish({ kind: 'stopped' }, true), ms);
		};
		worker.onmessage = (event) => {
			const message = event.data as CalculateResponse;
			switch (message.kind) {
				case 'ready':
					pool.markStarted(worker);
					return arm(limits.readMs);
				case 'read':
					return arm(limits.calculateMs + STOP_GRACE_MS);
				case 'done':
				case 'skipped':
					return finish(message);
				case 'stopped':
					return finish({ kind: 'stopped' });
				case 'error':
					// The page read the file, so it calculates it itself.
					return finish({ kind: 'unavailable' }, true);
			}
		};
		worker.onerror = (event) => {
			event.preventDefault?.();
			// The script never ran (a bundler without worker support, a blocked URL) or failed.
			finish({ kind: 'unavailable' }, true);
		};
		// A script that never starts (blocked without an error event) leaves the page to calculate.
		if (pool.hasStarted(worker)) arm(limits.readMs);
		else timer = setTimeout(() => finish({ kind: 'unavailable' }, true), limits.startMs);
		try {
			// `slice` copies the memory at once; `Uint8Array.from` would visit every byte.
			const copy = (bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).slice();
			const request: CalculateRequest = {
				bytes: copy.buffer,
				timeLimitMs: limits.calculateMs,
				...loadOptions(options),
			};
			worker.postMessage(request, [copy.buffer]);
		} catch {
			finish({ kind: 'unavailable' }, true);
		}
	});
}

/** Runs the calculation on the page within `bounds`; an error leaves the saved values. */
function calculateOnPage(workbook: Workbook, bounds: OpenCalculationLimits): OpenedWorkbook {
	try {
		const { ran, stopped } = calculateOnOpen(workbook, bounds);
		return { workbook, calculated: ran && !stopped, calculationStopped: !!stopped };
	} catch (error) {
		console.error(error);
		return { workbook, calculated: false, calculationStopped: false };
	}
}

/**
 * Opens workbooks on the page, calculating those that ask for it in a worker from `factory` (by
 * default the bundled `xlsx-open-worker.js`), or on the page where workers are unavailable.
 */
export function createWorkbookOpener(
	factory?: WorkerFactory,
	limits: OpenLimits = DEFAULT_OPEN_LIMITS,
): WorkbookOpener {
	const create =
		factory ??
		(typeof Worker === 'undefined'
			? undefined
			: () => new Worker(new URL('./xlsx-open-worker.js', import.meta.url), { type: 'module' }));
	const pool = create && new WorkerPool(create);
	const whole = { timeLimitMs: limits.calculateMs };
	return async (bytes, options) => {
		const { loadWorkbook } = await import('ooxml-core/xlsx/load');
		const started = performance.now();
		const workbook = await loadWorkbook(bytes, loadOptions(options));
		if (!needsCalculationOnOpen(workbook))
			return { workbook, calculated: false, calculationStopped: false };
		if (!pool) return calculateOnPage(workbook, whole);
		const first = calculateOnPage(workbook, {
			timeLimitMs: Math.max(limits.pageMs, performance.now() - started),
			workCap: limits.pageWorkCap,
		});
		if (!first.calculationStopped) return first;
		// What the page computed before it stopped is replaced by what the worker writes.
		const outcome = await calculateInWorker(pool, bytes, options, limits);
		switch (outcome.kind) {
			case 'done':
				applyCalculatedCells(workbook, outcome.cells);
				return { workbook, calculated: true, calculationStopped: false };
			case 'stopped':
				return { workbook, calculated: false, calculationStopped: true };
			default:
				return calculateOnPage(workbook, whole);
		}
	};
}
