import { hexToRgbChannels, toHex } from './color-primitives';
import { OFFICE_SIGMA_FACTORS } from './native-gradient-curve';
import { straightAlphaGradientStops, type StraightAlphaStop } from './straight-alpha-stops';

/** One sample every this many entries of the 256-entry sigma table (plus the last one). */
const SIGMA_SAMPLE_STEP = 8;

function round4(value: number): number {
	return Math.round(value * 10000) / 10000;
}

/**
 * Office's paint for a two-stop gradient whose stops sit at 0% and 100%, as
 * stops a plain sRGB interpolator can reproduce.
 *
 * Measured in PowerPoint 16 (COM `Slide.Export`): such a gradient is not a
 * linear sRGB blend. The colour weight follows the GDI+ sigma bell curve
 * (`OFFICE_SIGMA_FACTORS`) and the blend happens in gamma 2.2 light, so a
 * white to black ramp is already 186 (not 128) at the midpoint and 236 (not
 * 189) a quarter in. The same holds for linear (scaled or not), circle, rect
 * and shape paths, with or without alpha; alpha itself stays a straight linear
 * ramp. Gradients with three or more stops, or two stops that do not span the
 * full range, are blended plainly in sRGB per segment.
 *
 * Returns `undefined` when the stops are not such a pair, or when the colour
 * does not change (only alpha varies, which a plain interpolator already draws
 * correctly).
 */
export function officeTwoStopGradientStops(
	stops: readonly StraightAlphaStop[],
): StraightAlphaStop[] | undefined {
	const [from, to] = stops;
	if (stops.length !== 2 || !from || !to || from.position !== 0 || to.position !== 100) {
		return undefined;
	}
	const a = hexToRgbChannels(from.color);
	const b = hexToRgbChannels(to.color);
	if (!a || !b || (a.r === b.r && a.g === b.g && a.b === b.b)) return undefined;
	const fromAlpha = from.opacity ?? 1;
	const toAlpha = to.opacity ?? 1;
	const last = OFFICE_SIGMA_FACTORS.length - 1;
	const indices: number[] = [];
	for (let index = 0; index < last; index += SIGMA_SAMPLE_STEP) indices.push(index);
	indices.push(last);
	return indices.map((index) => {
		const t = index / last;
		const weight = (OFFICE_SIGMA_FACTORS[index] ?? 0) / 65536;
		const mix = (x: number, y: number) =>
			toHex(255 * ((x / 255) ** 2.2 * (1 - weight) + (y / 255) ** 2.2 * weight) ** (1 / 2.2));
		return {
			color: `#${mix(a.r, b.r)}${mix(a.g, b.g)}${mix(a.b, b.b)}`,
			position: round4(t * 100),
			...(fromAlpha === 1 && toAlpha === 1
				? {}
				: { opacity: round4(fromAlpha + (toAlpha - fromAlpha) * t) }),
		};
	});
}

/**
 * Stops that make a CSS gradient (premultiplied, plain sRGB) paint what Office
 * paints: the two-stop sigma profile when it applies, else the straight-alpha
 * re-sampling. The sigma samples are dense enough that premultiplied blending
 * between neighbours stays within a level of the straight ramp.
 */
export function officeGradientStops(stops: readonly StraightAlphaStop[]): StraightAlphaStop[] {
	return officeTwoStopGradientStops(stops) ?? straightAlphaGradientStops(stops);
}
