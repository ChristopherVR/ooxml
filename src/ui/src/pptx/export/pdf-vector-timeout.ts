/**
 * Bounds for the steps of a vector PDF export that wait on the browser (font loading, image
 * decoding, an isolated rasterisation that may inline web fonts). A step that never settles must
 * not hang the export: past its bound it yields the fallback, which every caller already treats as
 * a failed step (the element is left out and reported).
 */

/** How long one isolated rasterisation may take. */
export const RASTER_TIMEOUT_MS = 20_000;
/** How long one picture may take to load or decode. */
export const IMAGE_TIMEOUT_MS = 15_000;
/** How long to wait for the document's fonts before collecting anyway. */
export const FONTS_TIMEOUT_MS = 5_000;

/** `work`'s value, or `fallback` when it rejects or has not settled within `ms`. */
export function settleWithin<T, F>(work: PromiseLike<T>, ms: number, fallback: F): Promise<T | F> {
	return new Promise<T | F>((resolve) => {
		const timer = setTimeout(() => resolve(fallback), ms);
		Promise.resolve(work).then(
			(value) => {
				clearTimeout(timer);
				resolve(value);
			},
			() => {
				clearTimeout(timer);
				resolve(fallback);
			},
		);
	});
}
