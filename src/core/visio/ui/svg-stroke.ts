import type { VisioStyle } from '../model';

/**
 * Native SVG export substitutes 0.75 points for an exactly zero saved line weight.
 * Keep this output policy separate from source values and raster/device hairlines.
 * Nonzero weights, including weights below 0.75 points, remain literal.
 */
export function visioSvgStrokeStyle(style: VisioStyle): VisioStyle {
	return style.lineWidth === 0 ? { ...style, lineWidth: 0.75 / 72 } : style;
}
