/**
 * Opens workbook files off the page. A worker (`xlsx-open-worker.ts`) reads the file and runs the
 * calculation it asks for on open (`calcPr fullCalcOnLoad`), so a large file or slow formulas never
 * freeze the editor. The calculation has a hard time limit: past it the worker is ended and a fresh
 * one opens the file without calculating, so the workbook shows the values it was saved with.
 * Where workers are unavailable (no `Worker`, or the worker script cannot start) the file is
 * opened on the page instead, as before.
 */
import { calculateOnOpen, type Workbook } from 'ooxml-core/xlsx';

/** What the worker is asked to open. */
export interface OpenRequest {
	bytes: ArrayBuffer;
	fileName?: string;
	password?: string;
	/** Whether to run the calculation the file asks for on open. */
	calculate: boolean;
}

/** What the worker tells: it started, it read the file, its result, or why it failed. */
export type OpenResponse =
	| { kind: 'ready' }
	| { kind: 'read' }
	| { kind: 'done'; workbook: Workbook; calculated: boolean }
	| { kind: 'error'; message: string; code?: string };

export interface OpenOptions {
	fileName?: string;
	password?: string;
}

export interface OpenedWorkbook {
	workbook: Workbook;
	/** Whether the calculation the file asks for on open ran. */
	calculated: boolean;
	/** It ran past its time limit, so the workbook shows the values it was saved with. */
	calculationStopped: boolean;
}

export type WorkbookOpener = (
	bytes: Uint8Array | ArrayBuffer,
	options: OpenOptions,
) => Promise<OpenedWorkbook>;

export interface WorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: OpenRequest, transfer: Transferable[]): void;
	terminate(): void;
}
export type WorkerFactory = () => WorkerLike;

export interface OpenLimits {
	/** How long reading the file may take before the open fails. */
	readMs: number;
	/** How long the calculation on open may take before it is abandoned. */
	calculateMs: number;
}

export const DEFAULT_OPEN_LIMITS: OpenLimits = { readMs: 60_000, calculateMs: 10_000 };

/** The worker could not read the file within {@link OpenLimits.readMs}. */
export class OpenTimeoutError extends Error {
	constructor() {
		super('The workbook took too long to open.');
		this.name = 'OpenTimeoutError';
	}
}

class CalculationTimeout extends Error {}
class WorkerUnavailable extends Error {}

const loadOptions = (options: OpenOptions) => ({
	...(options.fileName === undefined ? {} : { fileName: options.fileName }),
	...(options.password === undefined ? {} : { password: options.password }),
});

/** Opens on the page: what the editor did before workers, and the fallback without them. */
export async function openOnPage(
	bytes: Uint8Array | ArrayBuffer,
	options: OpenOptions,
): Promise<OpenedWorkbook> {
	const { loadWorkbook } = await import('ooxml-core/xlsx/load');
	const workbook = await loadWorkbook(bytes, loadOptions(options));
	let calculated = false;
	try {
		calculated = calculateOnOpen(workbook);
	} catch (error) {
		console.error(error);
	}
	return { workbook, calculated, calculationStopped: false };
}

/** A failure the worker reported, rebuilt so callers see the loader's error and its `code`. */
function workerError(message: string, code: string | undefined): Error {
	const error = new Error(message);
	return code === undefined ? error : Object.assign(error, { code });
}

function runWorker(
	create: WorkerFactory,
	bytes: Uint8Array | ArrayBuffer,
	options: OpenOptions,
	calculate: boolean,
	limits: OpenLimits,
): Promise<{ workbook: Workbook; calculated: boolean }> {
	return new Promise((resolve, reject) => {
		let worker: WorkerLike;
		try {
			worker = create();
		} catch (cause) {
			reject(new WorkerUnavailable(String(cause)));
			return;
		}
		let started = false;
		let timer: ReturnType<typeof setTimeout> | undefined;
		const arm = (ms: number, error: () => Error) => {
			clearTimeout(timer);
			timer = setTimeout(() => finish(error()), ms);
		};
		const finish = (error?: Error, result?: { workbook: Workbook; calculated: boolean }) => {
			clearTimeout(timer);
			worker.onmessage = null;
			worker.onerror = null;
			worker.terminate();
			if (error) reject(error);
			else resolve(result!);
		};
		worker.onmessage = (event) => {
			const message = event.data as OpenResponse;
			switch (message.kind) {
				case 'ready':
					started = true;
					return;
				case 'read':
					arm(limits.calculateMs, () => new CalculationTimeout());
					return;
				case 'done':
					return finish(undefined, { workbook: message.workbook, calculated: message.calculated });
				case 'error':
					return finish(workerError(message.message, message.code));
			}
		};
		worker.onerror = (event) => {
			event.preventDefault?.();
			const reason = event.message || 'The workbook worker failed.';
			// Before `ready` the script never ran (a bundler without worker support, a blocked URL):
			// the page can open the file itself. Afterwards the open itself failed.
			finish(started ? new Error(reason) : new WorkerUnavailable(reason));
		};
		arm(limits.readMs, () => new OpenTimeoutError());
		try {
			const copy =
				bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0));
			const request: OpenRequest = { bytes: copy.buffer, calculate, ...loadOptions(options) };
			worker.postMessage(request, [copy.buffer]);
		} catch (cause) {
			finish(new WorkerUnavailable(String(cause)));
		}
	});
}

/**
 * Opens workbooks in a worker from `factory` (by default the bundled `xlsx-open-worker.js`), or on
 * the page where workers are unavailable.
 */
export function createWorkbookOpener(
	factory?: WorkerFactory,
	limits: OpenLimits = DEFAULT_OPEN_LIMITS,
): WorkbookOpener {
	if (!factory && typeof Worker === 'undefined') return openOnPage;
	const create =
		factory ??
		(() => new Worker(new URL('./xlsx-open-worker.js', import.meta.url), { type: 'module' }));
	return async (bytes, options) => {
		try {
			return {
				...(await runWorker(create, bytes, options, true, limits)),
				calculationStopped: false,
			};
		} catch (error) {
			if (error instanceof WorkerUnavailable) return openOnPage(bytes, options);
			if (!(error instanceof CalculationTimeout)) throw error;
		}
		const opened = await runWorker(create, bytes, options, false, limits);
		return { ...opened, calculationStopped: true };
	};
}
