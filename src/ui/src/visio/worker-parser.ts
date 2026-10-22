import { MAX_INPUT_BYTES } from 'ooxml-core/visio/ui';
import { loadVisio, type VisioDocument } from 'ooxml-core/visio';
export type CancellableParser = ((bytes: Uint8Array | ArrayBuffer) => Promise<VisioDocument>) & {
	cancel?: () => void;
	/**
	 * Reads only the pages `pageIds`: the document lists just those pages and leaves out what is
	 * not per page (see `ParseVsdxOptions.pages`). Merge it with `mergeVisioPages`.
	 */
	pages?: (bytes: Uint8Array | ArrayBuffer, pageIds: readonly string[]) => Promise<VisioDocument>;
	/** Releases the worker kept between page reads (and the styles, themes and masters it holds). */
	dispose?: () => void;
};
/** What a parse worker receives: the package, or the package and the pages to read. */
export type ParseWorkerRequest = ArrayBuffer | { bytes: ArrayBuffer; pages: string[] };
export interface WorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: ParseWorkerRequest, transfer: Transferable[]): void;
	terminate(): void;
}
export type WorkerFactory = () => WorkerLike;

/**
 * One disposable worker per document prevents old parses and long CPU tasks from occupying the UI.
 * Reading some pages of the drawing already open is the exception: the worker that did it stays
 * for the next such read (the shared drawing changes often), so it reads the styles, themes and
 * masters once while their parts stay the same. A whole read, a failure, a timeout, a cancel or
 * `dispose` ends it.
 */
export function createWorkerParser(factory?: WorkerFactory, timeoutMs = 15_000): CancellableParser {
	if (!factory && typeof Worker === 'undefined') return loadVisio;
	const create =
		factory ??
		(() => new Worker(new URL('./parse-worker.js', import.meta.url), { type: 'module' }));
	let cancel: (() => void) | undefined;
	let warm: WorkerLike | undefined;
	const release = (): void => {
		const worker = warm;
		warm = undefined;
		if (!worker) return;
		worker.onmessage = null;
		worker.onerror = null;
		worker.terminate();
	};
	const run = (
		bytes: Uint8Array | ArrayBuffer,
		pages: readonly string[] | undefined,
	): Promise<VisioDocument> => {
		cancel?.();
		if (!pages) release();
		if (bytes.byteLength > MAX_INPUT_BYTES)
			return Promise.reject(new Error('This viewer accepts files up to 32 MiB.'));
		return new Promise((resolve, reject) => {
			let worker: WorkerLike;
			try {
				worker = warm ?? create();
				warm = undefined;
			} catch (error) {
				reject(error);
				return;
			}
			let settled = false;
			const finish = (error?: Error, document?: VisioDocument) => {
				if (settled) return;
				settled = true;
				clearTimeout(timer);
				worker.onmessage = null;
				worker.onerror = null;
				cancel = undefined;
				if (!error && pages) warm = worker;
				else worker.terminate();
				if (error) reject(error);
				else resolve(document!);
			};
			const timer = setTimeout(
				() =>
					finish(
						new Error(
							`The diagram exceeded the ${timeoutMs / 1000}-second isolated parsing limit.`,
						),
					),
				timeoutMs,
			);
			cancel = () =>
				finish(new DOMException('The diagram load was superseded or cancelled.', 'AbortError'));
			worker.onmessage = (event) => {
				const result = event.data as { ok?: boolean; document?: VisioDocument; message?: string };
				if (result.ok && result.document) finish(undefined, result.document);
				else
					finish(
						new Error(result.message ?? 'The Visio parser worker returned an invalid response.'),
					);
			};
			worker.onerror = (event) =>
				finish(new Error(event.message || 'The Visio parser worker failed.'));
			try {
				const copy =
					bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0));
				worker.postMessage(pages ? { bytes: copy.buffer, pages: [...pages] } : copy.buffer, [
					copy.buffer,
				]);
			} catch (cause) {
				finish(cause instanceof Error ? cause : new Error(String(cause)));
			}
		});
	};
	const parse: CancellableParser = (bytes) => run(bytes, undefined);
	parse.pages = (bytes, pageIds) => run(bytes, pageIds);
	parse.cancel = () => cancel?.();
	parse.dispose = () => {
		cancel?.();
		release();
	};
	return parse;
}
