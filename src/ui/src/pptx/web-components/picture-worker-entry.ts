import {
	compressRaster,
	openRaster,
	rasterSegmentationSample,
	removeRasterBackground,
} from './picture-raster';
import type { PictureWorkerRequest, PictureWorkerResult } from './picture-worker-protocol';

let picture: Awaited<ReturnType<typeof openRaster>> | undefined;
let sample: ImageData | undefined;

async function respond(request: PictureWorkerRequest): Promise<PictureWorkerResult> {
	if (request.kind === 'open') {
		picture?.bitmap.close();
		picture = undefined;
		sample = undefined;
		picture = await openRaster(request.source, request.header);
		const { bitmap } = picture;
		return { kind: 'opened', width: bitmap.width, height: bitmap.height };
	}
	if (!picture) return { kind: 'error' };
	if (request.kind === 'compress') return compressRaster(picture, request.options);
	sample ??= rasterSegmentationSample(picture.bitmap);
	return removeRasterBackground(picture.bitmap, sample, request.marks, request.preview);
}

export function runPictureWorker(): void {
	// Importing the worker script outside a worker (server rendering, a bundler scan) does nothing.
	if (typeof globalThis.addEventListener !== 'function') return;
	globalThis.addEventListener('message', (event: MessageEvent<PictureWorkerRequest>) => {
		// Only the page that created this worker may send it pictures.
		if (event.origin && event.origin !== globalThis.location?.origin) return;
		void respond(event.data).then(
			(result) => globalThis.postMessage(result),
			() => globalThis.postMessage({ kind: 'error' }),
		);
	});
}
