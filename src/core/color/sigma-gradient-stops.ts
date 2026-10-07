import { hexToRgbChannels, toHex } from './color-primitives';
import { OFFICE_SIGMA_FACTORS } from './native-gradient-curve';

export interface SigmaGradientStop {
	offset: number;
	color: string;
	opacity?: number;
}

/** Native Office's measured opaque two-endpoint sigma/gamma-2.2 paint profile. */
export function sigmaGradientStops(
	stops: readonly SigmaGradientStop[],
): Array<SigmaGradientStop & { opacity: number }> | undefined {
	if (stops.length !== 2 || stops[0]?.offset !== 0 || stops[1]?.offset !== 1) return undefined;
	if (stops.some((stop) => (stop.opacity ?? 1) !== 1)) return undefined;
	const front = hexToRgbChannels(stops[0].color);
	const back = hexToRgbChannels(stops[1].color);
	if (!front || !back) return undefined;
	return OFFICE_SIGMA_FACTORS.map((value, index) => {
		const factor = value / 65536;
		const mix = (a: number, b: number) =>
			toHex(255 * ((a / 255) ** 2.2 * (1 - factor) + (b / 255) ** 2.2 * factor) ** (1 / 2.2));
		return {
			offset: Math.fround(index / 255),
			color: `#${mix(front.r, back.r)}${mix(front.g, back.g)}${mix(front.b, back.b)}`,
			opacity: 1,
		};
	});
}
