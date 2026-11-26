/// <reference lib="webworker" />
import { visioWorkerScope } from '../visio/worker-scope';
import { inspectRasterImage } from 'ooxml-core/drawingml';
import { compressRaster, openRaster, type PictureCompression } from '../form/picture-compression';

export interface PictureCompressRequest {
	bytes: Uint8Array<ArrayBuffer>;
	options: PictureCompression;
}

export type PictureCompressResponse =
	| { kind: 'compressed'; image: Blob }
	| { kind: 'unchanged' }
	| { kind: 'error' };

async function compress({
	bytes,
	options,
}: PictureCompressRequest): Promise<PictureCompressResponse> {
	const { mimeType, pixelWidth, pixelHeight } = inspectRasterImage(bytes);
	if (mimeType === 'image/gif') return { kind: 'error' };
	const picture = await openRaster(new Blob([bytes], { type: mimeType }), {
		width: pixelWidth,
		height: pixelHeight,
	});
	try {
		const { image } = await compressRaster(picture, options);
		return image.size < bytes.length ? { kind: 'compressed', image } : { kind: 'unchanged' };
	} finally {
		picture.bitmap.close();
	}
}

if (visioWorkerScope) {
	const worker = visioWorkerScope;
	worker.addEventListener(
		'message',
		(event: MessageEvent<PictureCompressRequest>) => {
			void compress(event.data).then(
				(result) => worker.postMessage(result),
				() => worker.postMessage({ kind: 'error' } satisfies PictureCompressResponse),
			);
		},
		{ once: true },
	);
}
