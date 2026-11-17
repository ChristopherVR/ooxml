/**
 * Convert a {@link RasterizeElementResult} (from `rasterize-element.ts`) into
 * a PNG `Blob`, handling both branches identically for every binding:
 * - `kind: 'canvas'` (the overwhelming majority of exports): wraps the
 *   canvas via `toBlob`.
 * - `kind: 'png-bytes'` (a tiled export whose full resolution exceeded the
 *   browser's canvas cap): wraps the already-encoded PNG bytes directly, no
 *   canvas involved at all, which is exactly what makes an export beyond the
 *   cap possible in the first place.
 *
 * Every binding's PNG-export / "copy slide as image" handler needs this
 * exact same branch, so it lives once here instead of being hand-copied five
 * times (the extraction trigger CLAUDE.md calls out: "I am making the same
 * edit in N bindings").
 */
import { releaseCanvas } from '../render/canvas-area-cap';
import type { RasterizeElementResult } from './rasterize-element';

/**
 * Convert a {@link RasterizeElementResult} to a PNG `Blob`. The result is consumed: its canvas is
 * released once encoded, since iOS refuses new canvases once their total memory is spent.
 */
export async function rasterResultToPngBlob(result: RasterizeElementResult): Promise<Blob> {
	if (result.kind === 'png-bytes') {
		return new Blob([result.bytes.slice()], { type: 'image/png' });
	}
	const canvas = result.canvas;
	return new Promise<Blob>((resolve, reject) => {
		canvas.toBlob((blob) => {
			releaseCanvas(canvas);
			if (blob) {
				resolve(blob);
			} else {
				reject(new Error('Canvas toBlob returned null'));
			}
		}, 'image/png');
	});
}

/** Convert a {@link RasterizeElementResult} to a PNG `data:` URL (for callers that need a string, not a `Blob`). */
export async function rasterResultToPngDataUrl(result: RasterizeElementResult): Promise<string> {
	if (result.kind === 'canvas') {
		const url = result.canvas.toDataURL('image/png');
		releaseCanvas(result.canvas);
		return url;
	}
	const blob = await rasterResultToPngBlob(result);
	return new Promise<string>((resolve, reject) => {
		const reader = new FileReader();
		reader.onloadend = () => resolve(reader.result as string);
		reader.onerror = reject;
		reader.readAsDataURL(blob);
	});
}
