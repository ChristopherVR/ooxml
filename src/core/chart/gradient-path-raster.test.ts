import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-path-profiles.json';
import paths from './__fixtures__/native-gradient-circle-shape-profiles.json';
import series from './__fixtures__/native-gradient-series-path-profiles.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../drawingml/drawing-fill';
import { drawingFillXml } from '../drawingml/write-fill';
import { resolveDrawingColor } from '../drawingml/drawing-color';
import { hexToRgbChannels } from '../color/color-primitives';
import { buildChartGradientDef, resolveChartGradient } from './gradient-definition';

for (const sample of [
	...reference.cases,
	...paths.cases.filter((sample) => sample.profile.includes('shape')),
	...series.cases
		.filter((sample) => sample.profile.includes('shape'))
		.map((sample) => ({
			...sample,
			width: sample.paintBounds.width,
			height: sample.paintBounds.height,
			samples: sample.samples.map((pixel) => ({
				...pixel,
				x: pixel.x - sample.paintBounds.x,
				y: pixel.y - sample.paintBounds.y,
			})),
		})),
])
	it(`paints native rectangular path gradient ${sample.name}`, () => {
		const read = (xml: string) =>
			parseDrawingFill(parseXml(`<a:spPr xmlns:a="${NS.a}">${xml}</a:spPr>`).documentElement)!;
		const fill = read(sample.fillXml);
		if (fill.kind !== 'gradient') throw new Error('Expected native gradient');
		const original = structuredClone(fill);
		const gradient = resolveChartGradient(fill, (color) => resolveDrawingColor(color));
		const def = buildChartGradientDef('native', gradient, {
			width: sample.width,
			height: sample.height,
			shape: 'rect',
		});
		expect(def.kind).toBe('rectPath');
		if (def.kind !== 'rectPath') throw new Error('Expected rectangular paint');
		const svg = parseXml(decodeURIComponent(def.href.split(',')[1]!));
		const mask = svg.getElementsByTagName('mask')[0];
		const rects = [
			...(svg.getElementsByTagName('g')[0] ?? svg).getElementsByTagName('rect'),
		].reverse();
		for (const pixel of sample.samples) {
			const x = (pixel.x / sample.width) * 100,
				y = (pixel.y / sample.height) * 100;
			const contains = (rect: Element) => {
				const left = Number(rect.getAttribute('x')),
					top = Number(rect.getAttribute('y'));
				return (
					x >= left &&
					y >= top &&
					x <= left + Number(rect.getAttribute('width')) &&
					y <= top + Number(rect.getAttribute('height'))
				);
			};
			const rect = rects.find(contains)!;
			const rgb = hexToRgbChannels(rect.getAttribute('fill')!)!;
			const alphaRect = mask && [...mask.getElementsByTagName('rect')].reverse().find(contains)!;
			const alpha = alphaRect ? hexToRgbChannels(alphaRect.getAttribute('fill')!)!.r : 255;
			expect(Math.abs(alpha - pixel.alpha)).toBeLessThanOrEqual(2);
			for (const [index, channel] of ['r', 'g', 'b'].entries())
				expect(
					Math.abs(
						(rgb[channel as keyof typeof rgb] * alpha) / 255 -
							(pixel.rgb[index]! * pixel.alpha) / 255,
					),
					`${pixel.x},${pixel.y} ${channel}`,
				).toBeLessThanOrEqual(pixel.alpha === 255 ? 2 : 3);
		}
		expect(fill).toEqual(original);
		expect(read(drawingFillXml(fill)!)).toMatchObject({
			stops: original.stops,
			path: original.path,
			fillToRect: original.fillToRect,
		});
	});
