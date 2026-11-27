/**
 * Browser pictures (data URLs, canvases, decoded images) as PDF image XObjects: 8-bit RGB
 * samples with an optional alpha plane, both Flate-compressed when the browser has
 * `CompressionStream`. Nothing is fetched: only `data:` and `blob:` sources are decoded.
 */
import type { PdfImageResource } from 'ooxml-core/pdf/write';

/** The longest side a decoded picture keeps (larger pictures are downscaled). */
export const PDF_MAX_IMAGE_SIDE = 2048;

async function deflate(bytes: Uint8Array, view: Window): Promise<Uint8Array | undefined> {
	const Compression = (view as unknown as { CompressionStream?: typeof CompressionStream })
		.CompressionStream;
	if (!Compression) return undefined;
	const stream = new Blob([Uint8Array.from(bytes)])
		.stream()
		.pipeThrough(new Compression('deflate'));
	return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** RGBA pixels as an image resource; the alpha plane is kept only when a pixel is translucent. */
export async function rgbaToPdfImage(
	rgba: Uint8ClampedArray,
	width: number,
	height: number,
	view: Window,
): Promise<PdfImageResource> {
	const rgb = new Uint8Array(width * height * 3);
	const alpha = new Uint8Array(width * height);
	let translucent = false;
	for (let pixel = 0; pixel < width * height; pixel++) {
		rgb[pixel * 3] = rgba[pixel * 4]!;
		rgb[pixel * 3 + 1] = rgba[pixel * 4 + 1]!;
		rgb[pixel * 3 + 2] = rgba[pixel * 4 + 2]!;
		alpha[pixel] = rgba[pixel * 4 + 3]!;
		if (alpha[pixel] !== 255) translucent = true;
	}
	const packed = await deflate(rgb, view);
	const packedAlpha = translucent ? await deflate(alpha, view) : undefined;
	return {
		width,
		height,
		data: packed ?? rgb,
		...(packed ? { filter: 'FlateDecode' as const } : {}),
		...(translucent
			? {
					alpha: {
						data: packedAlpha ?? alpha,
						...(packedAlpha ? { filter: 'FlateDecode' as const } : {}),
					},
				}
			: {}),
	};
}

/**
 * Draw `source` (an image, canvas or video frame) onto a canvas of at most `maxSide` pixels a
 * side and read it back. `undefined` when the pixels cannot be read (a tainted canvas).
 */
export async function drawableToPdfImage(
	doc: Document,
	source: CanvasImageSource,
	naturalWidth: number,
	naturalHeight: number,
	maxSide = PDF_MAX_IMAGE_SIDE,
): Promise<PdfImageResource | undefined> {
	const view = doc.defaultView;
	if (!view || !(naturalWidth > 0 && naturalHeight > 0)) return undefined;
	try {
		const scale = Math.min(1, maxSide / Math.max(naturalWidth, naturalHeight, 1));
		const width = Math.max(1, Math.round(naturalWidth * scale));
		const height = Math.max(1, Math.round(naturalHeight * scale));
		const canvas = doc.createElement('canvas');
		canvas.width = width;
		canvas.height = height;
		const context = canvas.getContext('2d');
		if (!context) return undefined;
		context.drawImage(source, 0, 0, width, height);
		const rgba = context.getImageData(0, 0, width, height).data;
		return await rgbaToPdfImage(rgba, width, height, view);
	} catch {
		return undefined;
	}
}

/** Decode a `data:` or `blob:` picture to RGB (and alpha) samples, downscaled to a bounded size. */
export async function decodePdfImage(
	doc: Document,
	href: string,
	maxSide = PDF_MAX_IMAGE_SIDE,
): Promise<PdfImageResource | undefined> {
	const view = doc.defaultView;
	if (!view || !(href.startsWith('data:image/') || href.startsWith('blob:'))) return undefined;
	try {
		const image = new view.Image();
		await new Promise<void>((resolve, reject) => {
			image.onload = () => resolve();
			image.onerror = () => reject(new Error('picture'));
			image.src = href;
		});
		return await drawableToPdfImage(doc, image, image.naturalWidth, image.naturalHeight, maxSide);
	} catch {
		return undefined;
	}
}
