import { safeColor, svgTile } from './svg-preview';
const num = (value: number) => String(Math.round(value * 100) / 100);
/** A horizontal strip of colour swatches (Change Colors rows). */
export function swatchStripSvg(
	colors: readonly string[],
	size: { width: number; height: number },
): string {
	const w = size.width / Math.max(colors.length, 1);
	const body = colors
		.map(
			(color, i) =>
				`<rect x="${num(i * w)}" y="0" width="${num(w)}" height="${size.height}" fill="${safeColor(color, '#808080')}"/>`,
		)
		.join('');
	return svgTile(size.width, size.height, '', body);
}
