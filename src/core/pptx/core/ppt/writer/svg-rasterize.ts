/**
 * SVG -> PNG rasteriser for the `.ppt` writer.
 *
 * A binary `.ppt` has no SVG BLIP, so an SVG picture needs a raster. When
 * the picture carries its own PNG fallback (`a:blip r:embed` next to
 * `asvg:svgBlip`) that fallback is embedded; but PowerPoint itself writes an
 * SVG picture with NO fallback (COM-measured: `Shapes.AddPicture` of an
 * `.svg`, then SaveAs `.pptx`, leaves a bare `<a:blip>` holding only the
 * `svgBlip` extension), and its own 97-2003 SaveAs then rasterises the SVG
 * to a PNG at twice the SVG's intrinsic size. This does the same through
 * the caller's rasteriser when the save passes one
 * (`PptxHandlerSaveOptions.rasterizeSvg`; Node.js hosts pass the one the
 * Node-only `pptx/node` entry exports), otherwise through the DOM's own SVG
 * renderer. With neither it returns `undefined`, where the caller degrades
 * to a placeholder with a warning. This module never loads a Node-only
 * package, so browser bundles of the engine stay free of them.
 *
 * @module ppt/writer/svg-rasterize
 */

import type { PptxSvgRasterizer } from '../../core/types';
import type { WPictureData } from './write-model';

/** Longest edge, in pixels, a rasterised SVG is clamped to. */
const MAX_EDGE_PX = 2048;

function hasDomRasteriser(): boolean {
	return (
		typeof document !== 'undefined' &&
		typeof Image !== 'undefined' &&
		typeof URL !== 'undefined' &&
		typeof URL.createObjectURL === 'function'
	);
}

/**
 * Raster size for an SVG: twice its intrinsic size (or the fallback size
 * when it declares none), clamped so the longest edge is at most 2048 px.
 */
export function svgRasterSize(
	intrinsicW: number,
	intrinsicH: number,
	fallbackW: number,
	fallbackH: number,
): { width: number; height: number } {
	const baseW = intrinsicW || fallbackW;
	const baseH = intrinsicH || fallbackH;
	const scale = Math.min(2, MAX_EDGE_PX / Math.max(baseW, baseH, 1));
	return {
		width: Math.max(1, Math.round(baseW * scale)),
		height: Math.max(1, Math.round(baseH * scale)),
	};
}

function loadImage(url: string): Promise<HTMLImageElement | undefined> {
	return new Promise((resolve) => {
		const img = new Image();
		img.onload = () => resolve(img);
		img.onerror = () => resolve(undefined);
		img.src = url;
	});
}

async function rasterizeSvgInDom(
	svg: Uint8Array,
	fallbackW: number,
	fallbackH: number,
): Promise<Uint8Array | undefined> {
	const url = URL.createObjectURL(new Blob([svg.slice()], { type: 'image/svg+xml' }));
	try {
		const img = await loadImage(url);
		if (!img) {
			return undefined;
		}
		const size = svgRasterSize(img.naturalWidth, img.naturalHeight, fallbackW, fallbackH);
		const canvas = document.createElement('canvas');
		canvas.width = size.width;
		canvas.height = size.height;
		const context = canvas.getContext('2d');
		if (!context) {
			return undefined;
		}
		context.drawImage(img, 0, 0, canvas.width, canvas.height);
		const blob = await new Promise<Blob | null>((resolve) => {
			canvas.toBlob(resolve, 'image/png');
		});
		return blob ? new Uint8Array(await blob.arrayBuffer()) : undefined;
	} catch {
		return undefined;
	} finally {
		URL.revokeObjectURL(url);
	}
}

/**
 * Rasterise SVG bytes to PNG, at twice the SVG's intrinsic size (or the
 * given fallback size when it declares none). `rasterize`, when given,
 * replaces the built-in DOM rasteriser. Returns `undefined` when the
 * rasteriser fails (returns nothing or throws), or with no rasteriser and
 * no DOM.
 */
export async function rasterizeSvg(
	svg: Uint8Array,
	fallbackWidthPx: number,
	fallbackHeightPx: number,
	rasterize?: PptxSvgRasterizer,
): Promise<WPictureData | undefined> {
	let png: Uint8Array | undefined;
	if (rasterize) {
		const sizeFor = (w: number, h: number) =>
			svgRasterSize(w, h, fallbackWidthPx, fallbackHeightPx);
		try {
			png = await rasterize(svg, sizeFor);
		} catch {
			png = undefined;
		}
	} else if (hasDomRasteriser()) {
		png = await rasterizeSvgInDom(svg, fallbackWidthPx, fallbackHeightPx);
	}
	return png ? { extension: 'png', bytes: png } : undefined;
}
