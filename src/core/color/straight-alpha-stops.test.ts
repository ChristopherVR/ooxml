import { describe, expect, it } from 'vitest';

import { hexToRgbChannels } from './color-primitives';
import { straightAlphaGradientStops } from './straight-alpha-stops';

/** Composite a premultiplied two-stop CSS blend at fraction `t` over `backdrop`. */
function premultipliedOver(
	from: { color: string; opacity?: number },
	to: { color: string; opacity?: number },
	t: number,
	backdrop: number,
): number[] {
	const a0 = from.opacity ?? 1;
	const a1 = to.opacity ?? 1;
	const c0 = hexToRgbChannels(from.color)!;
	const c1 = hexToRgbChannels(to.color)!;
	const alpha = a0 + (a1 - a0) * t;
	return (['r', 'g', 'b'] as const).map((k) => {
		const premul = c0[k] * a0 * (1 - t) + c1[k] * a1 * t;
		return premul + backdrop * (1 - alpha);
	});
}

/** Paint a stop list the way a premultiplied CSS gradient does, at `position` (0-100). */
function paint(
	stops: ReturnType<typeof straightAlphaGradientStops>,
	position: number,
	backdrop: number,
): number[] {
	const index = stops.findIndex((stop) => stop.position >= position);
	const to = stops[Math.max(1, index)]!;
	const from = stops[Math.max(1, index) - 1]!;
	const t = (position - from.position) / (to.position - from.position);
	return premultipliedOver(from, to, t, backdrop);
}

describe('straightAlphaGradientStops', () => {
	it('leaves stops alone when colour or opacity is constant', () => {
		const sameColour = [
			{ color: '#FFFFFF', position: 0 },
			{ color: '#FFFFFF', position: 95, opacity: 0 },
		];
		const opaque = [
			{ color: '#000000', position: 0 },
			{ color: '#FFFFFF', position: 100 },
		];
		expect(straightAlphaGradientStops(sameColour)).toStrictEqual(sameColour);
		expect(straightAlphaGradientStops(opaque)).toStrictEqual(opaque);
	});

	it('lets a transparent stop tint the ramp the way PowerPoint 16 paints it', () => {
		// Issue #141 slide 13: accent 00405E fading to 33667E at alpha 0 by 95%,
		// over a backdrop near 160. PowerPoint 16 paints (straight interpolation)
		// roughly 69/114/134 at a quarter of the ramp; premultiplied gives 59/106/128.
		const stops = straightAlphaGradientStops([
			{ color: '#00405E', position: 0 },
			{ color: '#33667E', position: 95, opacity: 0 },
		]);
		expect(stops).toHaveLength(9);
		expect(stops[4]).toStrictEqual({ color: '#1A536E', position: 47.5, opacity: 0.5 });
		for (const position of [10, 23.75, 40, 60, 80, 90]) {
			const t = position / 95;
			const alpha = 1 - t;
			const straight = [0x00, 0x40, 0x5e].map(
				(c, i) => (c + ([0x33, 0x66, 0x7e][i]! - c) * t) * alpha + 160 * (1 - alpha),
			);
			const painted = paint(stops, position, 160);
			painted.forEach((value, i) => expect(Math.abs(value - straight[i]!)).toBeLessThan(1));
		}
	});
});
