import { MAX_INPUT_BYTES } from 'ooxml-core/visio/ui';
import { VISIO_CLIPBOARD_MAX_CHARS } from 'ooxml-core/visio';

export interface ClipboardCaptureRequest {
	bytes: ArrayBuffer;
	pageId: string;
	shapeIds: readonly string[];
}
export interface ClipboardWorkerLike {
	onmessage: ((event: MessageEvent) => void) | null;
	onerror: ((event: ErrorEvent) => void) | null;
	postMessage(message: ClipboardCaptureRequest, transfer: Transferable[]): void;
	terminate(): void;
}
export type CancellableClipboardCapture = ((
	bytes: Uint8Array,
	pageId: string,
	shapeIds: readonly string[],
) => Promise<string>) & { cancel?: () => void };

/** Isolated source capture with an owned input, bounded response and cancellable lifetime. */
export function createWorkerClipboardCapture(
	factory?: () => ClipboardWorkerLike,
	timeoutMs = 20_000,
): CancellableClipboardCapture {
	const create =
		factory ??
		(() => {
			if (typeof Worker === 'undefined')
				throw new Error('Clipboard capture requires browser worker support.');
			return new Worker(new URL('./clipboard-worker.js', import.meta.url), { type: 'module' });
		});
	let cancel: (() => void) | undefined;
	const capture: CancellableClipboardCapture = async (bytes, pageId, shapeIds) => {
		cancel?.();
		if (
			bytes.byteLength > MAX_INPUT_BYTES ||
			typeof pageId !== 'string' ||
			!pageId ||
			pageId.length > 256 ||
			!Array.isArray(shapeIds) ||
			!shapeIds.length ||
			shapeIds.length > 1000
		)
			throw new Error('Invalid clipboard capture request.');
		const ids: string[] = [];
		for (let index = 0; index < shapeIds.length; index++) {
			const id = shapeIds[index];
			if (typeof id !== 'string' || !id || id.length > 1024)
				throw new Error('Invalid clipboard selection.');
			ids.push(id);
		}
		const copy = Uint8Array.from(bytes);
		return new Promise((resolve, reject) => {
			let worker: ClipboardWorkerLike;
			try {
				worker = create();
			} catch (error) {
				reject(error);
				return;
			}
			let settled = false;
			const finish = (error?: Error, text?: string) => {
				if (settled) return;
				settled = true;
				clearTimeout(timer);
				cancel = undefined;
				try {
					worker.onmessage = null;
					worker.onerror = null;
					worker.terminate();
				} catch (cause) {
					error ??= cause instanceof Error ? cause : new Error(String(cause));
				}
				if (error) reject(error);
				else resolve(text!);
			};
			const timer = setTimeout(
				() => finish(new Error('Clipboard capture exceeded the isolated time limit.')),
				timeoutMs,
			);
			cancel = () => finish(new DOMException('Clipboard capture was cancelled.', 'AbortError'));
			worker.onerror = (event) => finish(new Error(event.message || 'Clipboard capture failed.'));
			worker.onmessage = (event) => {
				const value = event.data;
				if (
					value?.ok === true &&
					typeof value.text === 'string' &&
					value.text.length <= VISIO_CLIPBOARD_MAX_CHARS
				)
					finish(undefined, value.text);
				else {
					const error = new Error(
						typeof value?.message === 'string' && value.message.length <= 4096
							? value.message
							: 'Invalid clipboard worker response.',
					);
					if (typeof value?.code === 'string' && value.code.length <= 128)
						Object.assign(error, { code: value.code });
					finish(error);
				}
			};
			try {
				worker.postMessage({ bytes: copy.buffer, pageId, shapeIds: ids }, [copy.buffer]);
			} catch (cause) {
				finish(cause instanceof Error ? cause : new Error(String(cause)));
			}
		});
	};
	capture.cancel = () => {
		cancel?.();
		cancel = undefined;
	};
	return capture;
}
