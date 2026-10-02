import { MAX_INPUT_BYTES } from './scene-validation.js';
import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
export type CancellableParser = ((bytes: Uint8Array | ArrayBuffer) => Promise<VisioDocument>) & {
	cancel?: () => void;
};
export interface WorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: ArrayBuffer, transfer: Transferable[]): void;
	terminate(): void;
}
export type WorkerFactory = () => WorkerLike;

/** One disposable worker per document prevents old parses and long CPU tasks from occupying the UI. */
export function createWorkerParser(factory?: WorkerFactory, timeoutMs = 15_000): CancellableParser {
	if (!factory && typeof Worker === 'undefined') return parseVsdx;
	const create =
		factory ??
		(() => new Worker(new URL('./parse-worker.js', import.meta.url), { type: 'module' }));
	let cancel: (() => void) | undefined;
	const parse: CancellableParser = (bytes) => {
		cancel?.();
		if (bytes.byteLength > MAX_INPUT_BYTES)
			return Promise.reject(new Error('This viewer accepts files up to 32 MiB.'));
		return new Promise((resolve, reject) => {
			let worker: WorkerLike;
			try {
				worker = create();
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
				worker.terminate();
				cancel = undefined;
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
				worker.postMessage(copy.buffer, [copy.buffer]);
			} catch (cause) {
				finish(cause instanceof Error ? cause : new Error(String(cause)));
			}
		});
	};
	parse.cancel = () => cancel?.();
	return parse;
}
