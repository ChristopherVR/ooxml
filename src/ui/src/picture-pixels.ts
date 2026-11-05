import type { DrawingPictureEffect } from 'ooxml-core/drawingml';
import { applyPictureHslPixels } from 'ooxml-core/drawingml';
import { pictureEffectFilter } from './picture-effects';
export interface PicturePixelImage {
	url: string;
	width: number;
	height: number;
}
/**
 * Processed pictures, newest last. Results are full-resolution PNG data URLs, often several MB, so
 * they get their own budget instead of the shared render cache's 8 MB cap, which turned them away:
 * an uncached result was recomputed (decode, recolour, PNG encode) on every render of the picture.
 */
const cache = new Map<string, PicturePixelImage>();
const CACHE_MAX_ENTRIES = 64;
const CACHE_MAX_CHARS = 32 * 1024 * 1024;
function remember(key: string, value: PicturePixelImage): void {
	cache.delete(key);
	if (value.url.length > CACHE_MAX_CHARS) return;
	cache.set(key, value);
	let chars = 0;
	for (const entry of cache.values()) chars += entry.url.length;
	for (const [oldest, entry] of cache) {
		if (cache.size <= CACHE_MAX_ENTRIES && chars <= CACHE_MAX_CHARS) break;
		chars -= entry.url.length;
		cache.delete(oldest);
	}
}
const pending = new Map<string, Promise<PicturePixelImage | undefined>>();
const keyOf = (
	url: string,
	effects: readonly DrawingPictureEffect[],
	scheme: Readonly<Record<string, string>>,
) => JSON.stringify([url, effects, scheme]);
export const needsPicturePixels = (effects: readonly DrawingPictureEffect[]) =>
	effects.some(
		(e) => e.kind === 'standard' && (e.name === 'hsl' || e.name === 'tint' || e.name === 'blur'),
	);
export function cachedPicturePixels(
	url: string,
	effects: readonly DrawingPictureEffect[],
	scheme: Readonly<Record<string, string>> = {},
): PicturePixelImage | undefined {
	const value = cache.get(keyOf(url, effects, scheme));
	return value ? { ...value } : undefined;
}
/** Rasterize nonlinear stages in order, flattening preceding SVG effects before each pixel stage. */
export function probePicturePixels(
	doc: Document,
	url: string,
	effects: readonly DrawingPictureEffect[],
	scheme: Readonly<Record<string, string>> = {},
): Promise<PicturePixelImage | undefined> {
	const key = keyOf(url, effects, scheme);
	const cached = cachedPicturePixels(url, effects, scheme);
	if (cached) return Promise.resolve(cached);
	const running = pending.get(key);
	if (running) return running;
	const load = (source: string) =>
		new Promise<HTMLImageElement>((resolve, reject) => {
			const image = doc.createElement('img');
			image.onload = () => resolve(image);
			image.onerror = () => reject(new Error('Picture decode failed'));
			image.src = source;
		});
	const task = (async () => {
		const source = await load(url);
		const canvas = doc.createElement('canvas');
		canvas.width = source.naturalWidth;
		canvas.height = source.naturalHeight;
		if (!canvas.width || !canvas.height || canvas.width * canvas.height > 16777216)
			return undefined;
		const pixelStage = (
			effect: DrawingPictureEffect,
		): effect is Parameters<typeof applyPictureHslPixels>[1] =>
			effect.kind === 'standard' && (effect.name === 'hsl' || effect.name === 'tint');
		// Pixel stages read the canvas back: keep it CPU-backed so those reads stay cheap.
		const context = canvas.getContext('2d', { willReadFrequently: effects.some(pixelStage) });
		if (!context) return undefined;
		context.drawImage(source, 0, 0);
		// Encoded only when an SVG stage or the result needs it; a PNG encode of a large picture is
		// the most expensive step here.
		let encoded: string | undefined;
		const dataUrl = () => (encoded ??= canvas.toDataURL('image/png'));
		let queue: DrawingPictureEffect[] = [];
		const flush = async () => {
			if (!queue.length) return;
			const ns = 'http://www.w3.org/2000/svg';
			const svg = doc.createElementNS(ns, 'svg');
			svg.setAttribute('xmlns', ns);
			svg.setAttribute('width', String(canvas.width));
			svg.setAttribute('height', String(canvas.height));
			const defs = doc.createElementNS(ns, 'defs');
			svg.append(defs);
			const result = pictureEffectFilter(doc, defs, queue, scheme);
			const image = doc.createElementNS(ns, 'image');
			image.setAttribute('href', dataUrl());
			image.setAttribute('width', String(canvas.width));
			image.setAttribute('height', String(canvas.height));
			if (result.filter) image.setAttribute('filter', result.filter);
			svg.append(image);
			const flattened = await load(`data:image/svg+xml,${encodeURIComponent(svg.outerHTML)}`);
			context.clearRect(0, 0, canvas.width, canvas.height);
			context.drawImage(flattened, 0, 0);
			encoded = undefined;
			queue = [];
		};
		for (const effect of effects) {
			if (pixelStage(effect)) {
				await flush();
				const image = context.getImageData(0, 0, canvas.width, canvas.height);
				applyPictureHslPixels(image.data, effect);
				context.putImageData(image, 0, 0);
				encoded = undefined;
			} else queue.push(effect);
		}
		await flush();
		const result = { url: dataUrl(), width: canvas.width, height: canvas.height };
		remember(key, result);
		return result;
	})()
		.catch(() => undefined)
		.finally(() => pending.delete(key));
	pending.set(key, task);
	return task;
}
