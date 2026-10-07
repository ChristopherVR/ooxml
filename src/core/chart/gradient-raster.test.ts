import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-raster.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../diagram/drawing-fill';
import { resolveDrawingColor } from '../diagram/drawing-color';
import { hexToRgbChannels } from '../color/color-primitives';
import { buildChartGradientDef, resolveChartGradient } from './gradient-definition';

for (const sample of reference.cases) {
	it(`matches native Excel raster samples for ${sample.name}`, () => {
		const fill = parseDrawingFill(
			parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
		)!;
		if (fill.kind !== 'gradient') throw new Error('Expected native gradient');
		const gradient = resolveChartGradient(fill, (color) => resolveDrawingColor(color));
		const def = buildChartGradientDef('native', gradient);
		if (def.kind !== 'linearGradient') throw new Error('Expected native linear gradient');
		expect(gradient.angle).toBe(sample.angle);
		expect(gradient.scaled).toBe(true);
		expect(gradient.stops).toHaveLength(2);
		expect(def.stops).toHaveLength(256);
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
			const right = def.stops.findIndex((stop) => stop.offset >= t);
			const a = def.stops[Math.max(0, right - 1)]!,
				b = def.stops[right]!;
			const ratio = a === b ? 0 : (t - a.offset) / (b.offset - a.offset);
			const front = hexToRgbChannels(a.color)!,
				back = hexToRgbChannels(b.color)!;
			for (const [index, channel] of ['r', 'g', 'b'].entries()) {
				const key = channel as keyof typeof front;
				const rendered = front[key] + (back[key] - front[key]) * ratio;
				expect(
					Math.abs(rendered - pixel.rgb[index]!),
					`${sample.name} at ${pixel.x},${pixel.y} channel ${channel}`,
				).toBeLessThanOrEqual(2);
			}
		}
	});
}
