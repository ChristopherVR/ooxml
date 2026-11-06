import { describe, expect, it } from 'vitest';

import { hexToRgbChannels } from './color-primitives';
import { officeGradientStops, officeTwoStopGradientStops } from './office-gradient-stops';

/** Paint a stop list as a plain sRGB interpolator does, at `position` (0-100). */
function paintRed(stops: Array<{ color: string; position: number }>, position: number): number {
	const index = stops.findIndex((stop) => stop.position >= position);
	if (index <= 0) return hexToRgbChannels(stops[0]!.color)!.r;
	const a = stops[index - 1]!;
	const b = stops[index]!;
	const t = (position - a.position) / (b.position - a.position);
	return hexToRgbChannels(a.color)!.r * (1 - t) + hexToRgbChannels(b.color)!.r * t;
}

describe('officeTwoStopGradientStops', () => {
	const whiteToBlack = [
		{ color: '#FFFFFF', position: 0 },
		{ color: '#000000', position: 100 },
	];

	it('paints the measured PowerPoint white to black ramp', () => {
		const stops = officeTwoStopGradientStops(whiteToBlack)!;
		// PowerPoint 16 exports 186 at the midpoint, 236 a quarter in and 108 at 74%.
		expect(Math.abs(paintRed(stops, 50) - 186)).toBeLessThan(2);
		expect(Math.abs(paintRed(stops, 26) - 236)).toBeLessThan(2.5);
		expect(Math.abs(paintRed(stops, 74) - 108)).toBeLessThan(2.5);
	});

	it('keeps alpha a straight linear ramp', () => {
		const stops = officeTwoStopGradientStops([
			{ color: '#FF0000', position: 0, opacity: 1 },
			{ color: '#0000FF', position: 100, opacity: 0 },
		])!;
		for (const stop of stops) expect(stop.opacity).toBeCloseTo(1 - stop.position / 100, 3);
	});

	it('leaves other stop lists to plain sRGB blending', () => {
		const three = [
			{ color: '#FF0000', position: 0 },
			{ color: '#00FF00', position: 50 },
			{ color: '#0000FF', position: 100 },
		];
		expect(officeTwoStopGradientStops(three)).toBeUndefined();
		expect(
			officeTwoStopGradientStops([
				{ color: '#FFFFFF', position: 20 },
				{ color: '#000000', position: 80 },
			]),
		).toBeUndefined();
		// Only alpha varies: a plain interpolator already paints it correctly.
		expect(
			officeTwoStopGradientStops([
				{ color: '#FF0000', position: 0, opacity: 1 },
				{ color: '#FF0000', position: 100, opacity: 0 },
			]),
		).toBeUndefined();
	});

	it('is what officeGradientStops returns for a two-stop pair', () => {
		expect(officeGradientStops(whiteToBlack).length).toBeGreaterThan(20);
		expect(officeGradientStops([{ color: '#FFFFFF', position: 0 }])).toHaveLength(1);
	});
});
