import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-raster.json';
import profiles from './__fixtures__/native-gradient-linear-profiles.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../diagram/drawing-fill';
import { drawingFillXml } from '../diagram/write-fill';
import { resolveDrawingColor } from '../diagram/drawing-color';
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
for (const sample of samples) {
	const knownGap =
		(sample.name.startsWith('three-transparent') &&
			!sample.name.includes('angle-180-') &&
			!sample.name.includes('angle-270-')) ||
		sample.name === 'coincident-angle-135-480x300';
	it(
		`${knownGap ? 'known native mismatch: ' : ''}matches native Excel raster samples for ${sample.name}`,
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
						((pixel.x / sample.width - def.x1) * dx + (pixel.y / sample.height - def.y1) * dy) /
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
