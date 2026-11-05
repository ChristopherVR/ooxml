import { hexToRgbChannels, toHex } from './color-primitives';

/** A gradient stop on a 0-100 position scale; `opacity` defaults to 1. */
export interface StraightAlphaStop {
	color: string;
	position: number;
	opacity?: number;
}

/** Sub-segments inserted between two stops whose colour and opacity both change. */
const STRAIGHT_ALPHA_STEPS = 8;

function round4(value: number): number {
	return Math.round(value * 10000) / 10000;
}

/**
 * Re-samples gradient stops so a renderer that interpolates in premultiplied
 * space (CSS gradients, per CSS Images 3) paints what Office paints.
 *
 * Office interpolates a gradient's colour and alpha independently (straight,
 * non-premultiplied), so the colour of a translucent or fully transparent stop
 * still tints the blend: measured in PowerPoint 16, `00405E` at 100% fading to
 * `33667E` at 0% alpha is 8 to 12 levels lighter mid-ramp than the same ramp
 * interpolated premultiplied, where the transparent stop's colour never shows.
 * Between every pair of stops whose colour and opacity both differ this inserts
 * {@link STRAIGHT_ALPHA_STEPS}-1 stops sampled from the straight interpolation;
 * the premultiplied blend between neighbouring samples then stays within a
 * level of the straight ramp. Stops that share a colour or an opacity
 * interpolate the same either way and are returned unchanged.
 */
export function straightAlphaGradientStops(
	stops: readonly StraightAlphaStop[],
): StraightAlphaStop[] {
	const out: StraightAlphaStop[] = [];
	stops.forEach((stop, index) => {
		const next = stops[index + 1];
		out.push(stop);
		if (!next) return;
		const fromAlpha = stop.opacity ?? 1;
		const toAlpha = next.opacity ?? 1;
		const from = hexToRgbChannels(stop.color);
		const to = hexToRgbChannels(next.color);
		if (fromAlpha === toAlpha || !from || !to || next.position <= stop.position) return;
		if (from.r === to.r && from.g === to.g && from.b === to.b) return;
		for (let step = 1; step < STRAIGHT_ALPHA_STEPS; step++) {
			const t = step / STRAIGHT_ALPHA_STEPS;
			const mix = (a: number, b: number) => toHex(a + (b - a) * t);
			out.push({
				color: `#${mix(from.r, to.r)}${mix(from.g, to.g)}${mix(from.b, to.b)}`,
				position: round4(stop.position + (next.position - stop.position) * t),
				opacity: round4(fromAlpha + (toAlpha - fromAlpha) * t),
			});
		}
	});
	return out;
}
