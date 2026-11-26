import { inkPathPoints } from './path';
import type { InkStroke } from './types';

/**
 * Ink strokes as an SVG document, for the views that paint drawings from markup (Word's page
 * layout and story previews, Excel's drawing layer). Strokes keep their own width when the
 * drawing is stretched; each draws round-capped at its brush colour and opacity.
 *
 * @module drawingml/inkml/svg
 */

const COLOR = /^#(?:[\da-f]{3}|[\da-f]{6})$/iu;
const PATH = /^[MLmlCcQqZz\d\s.,eE+-]*$/u;

/**
 * A polyline as a smooth curve: each sample becomes the control point of a quadratic from the
 * midpoint before it to the midpoint after it, so a stroke reads as drawn, not as joined segments.
 * Paths that are not plain polylines are kept as they are.
 */
export function smoothInkPath(path: string): string {
	const points = inkPathPoints(path);
	if (points.length < 3) return path;
	const r = (value: number) => Math.round(value * 100) / 100;
	const parts = [`M${r(points[0]!.x)} ${r(points[0]!.y)}`];
	for (let index = 1; index < points.length - 1; index++) {
		const point = points[index]!;
		const next = points[index + 1]!;
		parts.push(
			`Q${r(point.x)} ${r(point.y)} ${r((point.x + next.x) / 2)} ${r((point.y + next.y) / 2)}`,
		);
	}
	const last = points.at(-1)!;
	parts.push(`L${r(last.x)} ${r(last.y)}`);
	return parts.join('');
}

function finite(value: number, fallback: number): number {
	return Number.isFinite(value) && value > 0 ? value : fallback;
}

/**
 * `strokes` (drawn in a `viewWidth` x `viewHeight` space, by default the box itself) stretched
 * onto a `width` x `height` box. Paths and colours that are not plain path data or hex colours are
 * skipped, so nothing from the part reaches the markup unchecked.
 */
export function inkSvg(
	strokes: readonly InkStroke[],
	width: number,
	height: number,
	viewWidth = width,
	viewHeight = height,
): string {
	const w = finite(width, 1);
	const h = finite(height, 1);
	const scaled = viewWidth !== width || viewHeight !== height;
	const paths = strokes
		.filter((stroke) => stroke.path && PATH.test(stroke.path))
		.map((stroke) => {
			const color = COLOR.test(stroke.color) ? stroke.color : '#000000';
			const opacity = Math.max(
				0,
				Math.min(1, Number.isFinite(stroke.opacity) ? stroke.opacity : 1),
			);
			return (
				`<path d="${smoothInkPath(stroke.path)}" fill="none" stroke="${color}" stroke-width="${finite(stroke.width, 1)}"` +
				` stroke-linecap="round" stroke-linejoin="round"` +
				(opacity < 1 ? ` stroke-opacity="${opacity}"` : '') +
				(scaled ? ' vector-effect="non-scaling-stroke"' : '') +
				'/>'
			);
		});
	return (
		`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"` +
		` viewBox="0 0 ${finite(viewWidth, w)} ${finite(viewHeight, h)}" preserveAspectRatio="none"` +
		` overflow="visible">${paths.join('')}</svg>`
	);
}
