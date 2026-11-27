/**
 * A blurred silhouette for soft shadows in a vector PDF. PDF has no blur, so the shapes are
 * rasterised (anti-aliased scanlines for fills, distance to the outline for strokes) at a
 * resolution set by the blur, blurred with a Gaussian (three box passes) and written as one
 * small image of the shadow colour whose soft mask (`/SMask`) is the blurred coverage. Only the
 * shadow is a picture; the shapes and text above it stay vectors. DOM-free.
 */
import type { PathSegment } from './pdf-path';
import {
	fillCoverage,
	flattenSegments,
	gaussianBlur,
	IDENTITY,
	strokeCoverage,
	transformed,
	type Point,
} from './pdf-raster';
import type { PdfImageResource, PdfVectorPage } from './pdf-vector';

/** One shape of a silhouette, in its own units mapped by `matrix` (SVG order, a b c d e f). */
export interface SilhouettePart {
	segments: readonly PathSegment[];
	/** Fill rule when the part is filled. */
	fill?: 'nonzero' | 'evenodd';
	/** Stroke width when the part is stroked, in the part's units. */
	stroke?: number;
	matrix?: readonly number[];
}

export interface BlurredSilhouette {
	resource: PdfImageResource;
	/** The image's box in silhouette units (y down: `y` is its top edge). */
	x: number;
	y: number;
	width: number;
	height: number;
}

/** Longest side of the raster in pixels: a blurred edge needs few. */
const MAX_SIDE = 384;

/**
 * The parts' silhouette blurred with a Gaussian of deviation `sigma` (silhouette units), as an
 * image of `color` (0..1 channels) with the blurred coverage as its alpha; `undefined` when the
 * parts cover nothing.
 */
export function blurredSilhouette(
	parts: readonly SilhouettePart[],
	sigma: number,
	color: { r: number; g: number; b: number },
): BlurredSilhouette | undefined {
	let minX = Infinity,
		minY = Infinity,
		maxX = -Infinity,
		maxY = -Infinity,
		reach = 0;
	const scales = parts.map((part) => {
		const m = part.matrix ?? IDENTITY;
		return Math.sqrt(Math.abs(m[0]! * m[3]! - m[1]! * m[2]!));
	});
	parts.forEach((part, index) => {
		if (part.stroke) reach = Math.max(reach, (part.stroke * scales[index]!) / 2);
		for (const segment of part.segments) {
			if (segment.t === 'Z') continue;
			const coordinates =
				segment.t === 'C'
					? [segment.x1, segment.y1, segment.x2, segment.y2, segment.x, segment.y]
					: [segment.x, segment.y];
			for (let i = 0; i < coordinates.length; i += 2) {
				const [x, y] = transformed(part.matrix ?? IDENTITY, coordinates[i]!, coordinates[i + 1]!);
				minX = Math.min(minX, x);
				minY = Math.min(minY, y);
				maxX = Math.max(maxX, x);
				maxY = Math.max(maxY, y);
			}
		}
	});
	if (!Number.isFinite(minX + minY + maxX + maxY) || !(sigma >= 0)) return undefined;
	const pad = 3 * sigma + reach;
	const left = minX - pad;
	const top = minY - pad;
	const spanX = maxX - minX + 2 * pad;
	const spanY = maxY - minY + 2 * pad;
	const pixel = Math.max(sigma / 3, Math.max(spanX, spanY) / MAX_SIDE, 1e-6);
	const width = Math.max(1, Math.ceil(spanX / pixel));
	const height = Math.max(1, Math.ceil(spanY / pixel));
	// Silhouette units to raster pixels.
	const toPixels = [1 / pixel, 0, 0, 1 / pixel, -left / pixel, -top / pixel];
	const fills = new Float32Array(width * height);
	const strokes = new Float32Array(width * height);
	const nonzero: Point[][] = [];
	parts.forEach((part, index) => {
		const m = part.matrix ?? IDENTITY;
		const matrix = [
			m[0]! * toPixels[0]!,
			m[1]! * toPixels[3]!,
			m[2]! * toPixels[0]!,
			m[3]! * toPixels[3]!,
			m[4]! * toPixels[0]! + toPixels[4]!,
			m[5]! * toPixels[3]! + toPixels[5]!,
		];
		const lines = flattenSegments(part.segments, matrix, 0.5);
		if (part.fill === 'evenodd')
			fillCoverage(
				fills,
				width,
				height,
				lines.map((line) => line.points),
				true,
			);
		else if (part.fill) nonzero.push(...lines.map((line) => line.points));
		if (part.stroke)
			for (const line of lines)
				strokeCoverage(strokes, width, height, line, (part.stroke * scales[index]!) / 2 / pixel);
	});
	if (nonzero.length) fillCoverage(fills, width, height, nonzero, false);
	let any = false;
	for (let index = 0; index < fills.length; index++) {
		const fill = Math.min(1, fills[index]!);
		fills[index] = 1 - (1 - fill) * (1 - strokes[index]!);
		if (fills[index]! > 0) any = true;
	}
	if (!any) return undefined;
	gaussianBlur(fills, width, height, sigma / pixel);
	const alpha = new Uint8Array(width * height);
	for (let index = 0; index < alpha.length; index++)
		alpha[index] = Math.round(Math.max(0, Math.min(1, fills[index]!)) * 255);
	const rgb = new Uint8Array(width * height * 3);
	const channels = [color.r, color.g, color.b].map((c) =>
		Math.round(Math.max(0, Math.min(1, c)) * 255),
	);
	for (let index = 0; index < rgb.length; index++) rgb[index] = channels[index % 3]!;
	return {
		resource: { width, height, data: rgb, alpha: { data: alpha } },
		x: left,
		y: top,
		width: width * pixel,
		height: height * pixel,
	};
}

/**
 * Paints the blurred silhouette of `parts` at `opacity` in the current space (y down). Returns
 * false when there was nothing to paint.
 */
export function drawBlurredSilhouette(
	page: PdfVectorPage,
	parts: readonly SilhouettePart[],
	sigma: number,
	color: { r: number; g: number; b: number },
	opacity: number,
): boolean {
	if (!(opacity > 0)) return false;
	const mask = blurredSilhouette(parts, sigma, color);
	if (!mask) return false;
	page.save();
	if (opacity < 1) page.alpha(Math.min(1, opacity));
	page.concat([mask.width, 0, 0, -mask.height, mask.x, mask.y + mask.height]);
	page.drawImage(mask.resource);
	page.restore();
	return true;
}
