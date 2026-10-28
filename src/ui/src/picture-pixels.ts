import type { DrawingPictureEffect } from 'ooxml-core/drawingml';
import { applyPictureHslPixels } from 'ooxml-core/drawingml';
import { pictureEffectFilter } from './picture-effects';
import { setBoundedRenderCache } from './render-cache';
export interface PicturePixelImage {
	url: string;
	width: number;
	height: number;
}
const cache = new Map<string, string>();
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
	return value ? (JSON.parse(value) as PicturePixelImage) : undefined;
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
		const context = canvas.getContext('2d');
		if (!context) return undefined;
		context.drawImage(source, 0, 0);
		let dataUrl = canvas.toDataURL('image/png');
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
			image.setAttribute('href', dataUrl);
			image.setAttribute('width', String(canvas.width));
			image.setAttribute('height', String(canvas.height));
			if (result.filter) image.setAttribute('filter', result.filter);
			svg.append(image);
			const flattened = await load(`data:image/svg+xml,${encodeURIComponent(svg.outerHTML)}`);
			context.clearRect(0, 0, canvas.width, canvas.height);
			context.drawImage(flattened, 0, 0);
			dataUrl = canvas.toDataURL('image/png');
			queue = [];
		};
		for (const effect of effects) {
			if (effect.kind === 'standard' && (effect.name === 'hsl' || effect.name === 'tint')) {
				await flush();
				const image = context.getImageData(0, 0, canvas.width, canvas.height);
				applyPictureHslPixels(image.data, effect);
				context.putImageData(image, 0, 0);
				dataUrl = canvas.toDataURL('image/png');
			} else queue.push(effect);
		}
		await flush();
		const result = { url: dataUrl, width: canvas.width, height: canvas.height };
		setBoundedRenderCache(cache, key, JSON.stringify(result), 64);
		return result;
	})()
		.catch(() => undefined)
		.finally(() => pending.delete(key));
	pending.set(key, task);
	return task;
}
