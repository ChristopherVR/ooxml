import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-raster.json';
import profiles from './__fixtures__/native-gradient-linear-profiles.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../drawingml/drawing-fill';
import { drawingFillXml } from '../drawingml/write-fill';
import { resolveDrawingColor } from '../drawingml/drawing-color';
import { hexToRgbChannels } from '../color/color-primitives';
import { buildChartGradientDef, resolveChartGradient } from './gradient-definition';

const samples = [
	...reference.cases.map((sample) => ({
		...sample,
		paintStops: 256,
		samples: sample.samples.map((pixel) => ({ ...pixel, alpha: 255 })),
	})),
	...profiles.cases.map((sample) => ({
		...sample,
		paintStops: sample.profile.startsWith('three') ? 3 : 2,
	})),
];
it('preserves native stop insertion order, opacity and geometry across every profile save', () => {
	for (const sample of samples) {
		const read = (xml: string) =>
			parseDrawingFill(parseXml(`<a:spPr xmlns:a="${NS.a}">${xml}</a:spPr>`).documentElement)!;
		const fill = read(sample.fillXml);
		if (fill.kind !== 'gradient') throw new Error('Expected gradient');
		const source = structuredClone(fill);
		const gradient = resolveChartGradient(fill, (color) => resolveDrawingColor(color));
		const def = buildChartGradientDef('native', gradient);
		expect(gradient.stops, sample.name).toHaveLength(fill.stops.length);
		expect(def.stops, sample.name).toHaveLength(sample.paintStops);
		expect(def.kind, sample.name).toBe('linearGradient');
		expect(gradient.angle, sample.name).toBe(sample.angle);
		expect(fill, sample.name).toEqual(source);
		expect(read(drawingFillXml(fill)!), sample.name).toMatchObject({
			stops: source.stops,
			angle: source.angle,
			scaled: source.scaled,
		});
	}
});
/**
 * Accepted native mismatches, named so each one carries its reason. `fails` keeps the
 * acceptance honest: if the renderer ever matches the reference, the case reports as
 * unexpectedly passing and the entry must be deleted.
 */
const acceptedNativeGaps = [
	{
		id: 'translucent-three-stop-colour-weighting',
		// Angles 180 and 270 peak at 2.40 to 2.74 levels and pass; the other eight angles at
		// both sizes of the three-transparent profile (16 cases) miss.
		matches: (name: string) =>
			name.startsWith('three-transparent') &&
			!name.includes('angle-180-') &&
			!name.includes('angle-270-'),
		reason:
			'Sixteen translucent three-stop cases miss by 3.22 to 3.67 premultiplied levels ' +
			'(limit 3); alpha stays within 0.76. Native red in the 0-56% segment sits about 2.3 ' +
			'levels below the model and green about 0.9 above, i.e. colour leans about 0.01 of ' +
			'the segment toward the more opaque stop (premultiplied interpolation would lean ' +
			'about 0.08), plus roughly one level of PARGB-to-ARGB truncation in the reference ' +
			'export (opaque green reads 254). An alpha weight of a^0.05 would pass but is an ' +
			'unexplained fit, so it is not applied and the global tolerance is not loosened. ' +
			'Needed evidence: native captures of a two-stop red-to-green fill with unequal ' +
			'alphas, and of the same fill with its alphas swapped, to separate the colour ' +
			'weighting from export truncation.',
	},
];
for (const sample of samples) {
	const accepted = acceptedNativeGaps.find((gap) => gap.matches(sample.name));
	const knownGap = accepted !== undefined;
	it(
		`${accepted ? `accepted native gap (${accepted.id}): ` : ''}matches native Excel raster samples for ${sample.name}`,
		{ fails: knownGap },
		() => {
			const fill = parseDrawingFill(
				parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
			)!;
			if (fill.kind !== 'gradient') throw new Error('Expected native gradient');
			const gradient = resolveChartGradient(fill, (color) => resolveDrawingColor(color));
			const def = buildChartGradientDef('native', gradient);
			if (def.kind !== 'linearGradient') throw new Error('Expected native linear gradient');
			expect(gradient.angle).toBe(sample.angle);
			expect(gradient.scaled).toBe(true);
			expect(gradient.stops).toHaveLength(fill.stops.length);
			expect(def.stops).toHaveLength(sample.paintStops);
			const dx = def.x2 - def.x1,
				dy = def.y2 - def.y1;
			for (const pixel of sample.samples) {
				const t = Math.min(
					1,
					Math.max(
						0,
						(((pixel.x + 0.5) / sample.width - def.x1) * dx +
							((pixel.y + 0.5) / sample.height - def.y1) * dy) /
							(dx * dx + dy * dy),
					),
				);
				const found = def.stops.findIndex((stop) => stop.offset >= t);
				const right = found < 0 ? def.stops.length - 1 : found;
				const a = def.stops[t >= def.stops.at(-1)!.offset ? right : Math.max(0, right - 1)]!,
					b = def.stops[right]!;
				const ratio =
					a === b || a.offset === b.offset
						? 0
						: Math.min(1, Math.max(0, (t - a.offset) / (b.offset - a.offset)));
				const front = hexToRgbChannels(a.color)!,
					back = hexToRgbChannels(b.color)!;
				const alpha = (a.opacity ?? 1) + ((b.opacity ?? 1) - (a.opacity ?? 1)) * ratio;
				expect(Math.abs(alpha * 255 - pixel.alpha)).toBeLessThanOrEqual(2);
				for (const [index, channel] of ['r', 'g', 'b'].entries()) {
					const key = channel as keyof typeof front;
					const rendered = front[key] + (back[key] - front[key]) * ratio;
					expect(
						Math.abs(rendered * alpha - (pixel.rgb[index]! * pixel.alpha) / 255),
						`${sample.name} at ${pixel.x},${pixel.y} channel ${channel}`,
					).toBeLessThanOrEqual(pixel.alpha === 255 ? 2 : 3);
				}
			}
		},
	);
}
