import { MAX_PICTURE_EDIT_BYTES, readPictureRasterHeader } from 'ooxml-core/pptx';
import type { PictureWorkerRequest, PictureWorkerResult } from './picture-worker-protocol';

/** Only local URLs minted for file content are accepted. */
export async function localPictureBlob(url: string, signal: AbortSignal): Promise<Blob> {
	const localBlob = url.startsWith('blob:');
	const localData =
		url.length <= Math.ceil(MAX_PICTURE_EDIT_BYTES / 3) * 4 + 32 &&
		/^data:image\/(?:png|jpe?g);base64,/iu.test(url);
	if (!localBlob && !localData) throw new Error('Picture source is not a local raster');
	const response = await fetch(url, { signal });
	if (!response.ok) throw new Error('Picture is unavailable');
	const blob = await response.blob();
	if (blob.size > MAX_PICTURE_EDIT_BYTES) throw new Error('Picture exceeds the byte limit');
	return blob;
}

export function pictureDataUrl(blob: Blob): Promise<string> {
	if (blob.size > MAX_PICTURE_EDIT_BYTES)
		return Promise.reject(new Error('Picture exceeds the byte limit'));
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = () =>
			typeof reader.result === 'string' ? resolve(reader.result) : reject(reader.error);
		reader.onerror = () => reject(reader.error);
		reader.readAsDataURL(blob);
	});
}

export function createPictureWorker(signal: AbortSignal) {
	const worker = new Worker(new URL('./picture-worker.mjs', import.meta.url), { type: 'module' });
	let closed = false;
	const close = () => {
		closed = true;
		worker.terminate();
	};
	const send = (request: PictureWorkerRequest): Promise<PictureWorkerResult> =>
		new Promise((resolve, reject) => {
			signal.throwIfAborted();
			if (closed) {
				reject(new Error('Picture processing is unavailable'));
				return;
			}
			const end = (result?: PictureWorkerResult) => {
				clearTimeout(timeout);
				worker.onmessage = null;
				worker.onerror = null;
				signal.removeEventListener('abort', abort);
				if (result && result.kind !== 'error') resolve(result);
				else reject(new Error('Cannot process picture'));
			};
			const abort = () => {
				close();
				end();
			};
			const timeout = setTimeout(abort, 30_000);
			signal.addEventListener('abort', abort, { once: true });
			worker.onmessage = (event: MessageEvent<PictureWorkerResult>) => end(event.data);
			worker.onerror = abort;
			worker.postMessage(request);
		});
	return {
		send,
		async open(blob: Blob) {
			if (blob.size > MAX_PICTURE_EDIT_BYTES) throw new Error('Picture exceeds the byte limit');
			const bytes = await blob.arrayBuffer();
			const header = readPictureRasterHeader(bytes);
			if (!header) throw new Error('Picture type or dimensions are unsupported');
			const source = blob.slice(0, blob.size, header.type);
			const result = await send({ kind: 'open', source, header });
			if (result.kind !== 'opened') throw new Error('Cannot open picture');
			return { ...result, source };
		},
		close,
	};
}
