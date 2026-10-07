import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-circle-shape-profiles.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../diagram/drawing-fill';
import { drawingFillXml } from '../diagram/write-fill';
import { resolveDrawingColor } from '../diagram/drawing-color';
import { hexToRgbChannels } from '../color/color-primitives';
import { buildChartGradientDef, resolveChartGradient } from './gradient-definition';
import { chartGradientMarkup } from './gradient-markup';

for (const sample of reference.cases.filter((sample) => sample.profile.includes('circle')))
	it(`paints physical circular native gradient ${sample.name}`, () => {
		const read = (xml: string) =>
			parseDrawingFill(parseXml(`<a:spPr xmlns:a="${NS.a}">${xml}</a:spPr>`).documentElement)!;
		const fill = read(sample.fillXml);
		if (fill.kind !== 'gradient') throw new Error('Expected native gradient');
		const original = structuredClone(fill);
		const def = buildChartGradientDef(
			'native',
			resolveChartGradient(fill, (color) => resolveDrawingColor(color)),
			{ width: sample.width, height: sample.height },
		);
		if (def.kind !== 'radialGradient') throw new Error('Expected circular paint');
		const aspect = sample.height / sample.width;
		for (const pixel of sample.samples) {
			const t = Math.min(
				1,
				Math.hypot(pixel.x / sample.width - def.cx, (pixel.y / sample.height - def.cy) * aspect) /
					def.r,
			);
			const right = Math.max(
				1,
				def.stops.findIndex((stop) => stop.offset >= t),
			);
			const a = def.stops[right - 1]!,
				b = def.stops[right]!;
			const ratio = (t - a.offset) / (b.offset - a.offset);
			const front = hexToRgbChannels(a.color)!,
				back = hexToRgbChannels(b.color)!;
			const alpha = (a.opacity ?? 1) + ((b.opacity ?? 1) - (a.opacity ?? 1)) * ratio;
			expect(Math.abs(alpha * 255 - pixel.alpha)).toBeLessThanOrEqual(2);
			for (const [index, channel] of ['r', 'g', 'b'].entries()) {
				const key = channel as keyof typeof front;
				const rendered = front[key] + (back[key] - front[key]) * ratio;
				expect(
					Math.abs(rendered * alpha - (pixel.rgb[index]! * pixel.alpha) / 255),
					`${pixel.x},${pixel.y} ${channel}`,
				).toBeLessThanOrEqual(pixel.alpha === 255 ? 2 : 3);
			}
		}
		const markup = parseXml(chartGradientMarkup(def)).documentElement;
		expect(markup.getAttribute('gradientTransform')).toBe(def.gradientTransform ?? null);
		expect(fill).toEqual(original);
		expect(read(drawingFillXml(fill)!)).toMatchObject({
			stops: original.stops,
			path: original.path,
			fillToRect: original.fillToRect,
			tileRect: original.tileRect,
		});
	});
