import { RASTER_IMAGE_LIMITS } from 'ooxml-core/drawingml';
import type { PictureCompression } from '../form/picture-compression';
import type {
	PictureCompressRequest,
	PictureCompressResponse,
} from './xlsx-picture-compress-worker';

/** Compress a copy so the workbook keeps its original media available for undo. */
export async function compressPicture(
	bytes: Uint8Array,
	options: PictureCompression,
): Promise<Blob | undefined> {
	if (bytes.length > RASTER_IMAGE_LIMITS.maxImageBytes) throw new RangeError('Picture too large');
	const copy = new Uint8Array(bytes);
	const worker = new Worker(new URL('./xlsx-picture-compress-worker.js', import.meta.url), {
		type: 'module',
	});
	let timeout: ReturnType<typeof setTimeout> | undefined;
	try {
		return await new Promise((resolve, reject) => {
			const fail = () => reject(new Error('Picture compression failed'));
			timeout = setTimeout(fail, 30_000);
			worker.onerror = fail;
			worker.onmessageerror = fail;
			worker.onmessage = ({ data }: MessageEvent<PictureCompressResponse>) => {
				if (data.kind === 'compressed') resolve(data.image);
				else if (data.kind === 'unchanged') resolve(undefined);
				else fail();
			};
			worker.postMessage({ bytes: copy, options } satisfies PictureCompressRequest, [copy.buffer]);
		});
	} finally {
		clearTimeout(timeout);
		worker.terminate();
	}
}
