import { expect, it } from 'vitest';
import reference from './__fixtures__/native-gradient-three-stop-backdrops.json';
import { parseXml, NS } from '../xml';
import { parseDrawingFill } from '../drawingml/drawing-fill';
import { resolveDrawingColor } from '../drawingml/drawing-color';
import { hexToRgbChannels } from '../color/color-primitives';
import { buildChartGradientDef, resolveChartGradient } from './gradient-definition';

// Excel 16.0 build 20430 series fills (red 63% / green 87% at 56% / white) painted over an opaque
// black chart area and over an opaque white one (`-Backdrop` in the recorder). Opaque exports
// carry no PARGB-to-ARGB truncation, so the per-pixel alpha and colour solved from the pair are
// exact to rounding, which separates the capture method from the renderer.
type Pixel = { x: number; y: number; onBlack: number[]; onWhite: number[] };

function modelAt(sample: (typeof reference.cases)[number], pixel: Pixel) {
	const fill = parseDrawingFill(
		parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
	)!;
	if (fill.kind !== 'gradient') throw new Error('Expected gradient');
	const def = buildChartGradientDef(
		'native',
		resolveChartGradient(fill, (color) => resolveDrawingColor(color)),
	);
	if (def.kind !== 'linearGradient') throw new Error('Expected linear gradient');
	const { x, y, width, height } = sample.paintBounds;
	const dx = def.x2 - def.x1,
		dy = def.y2 - def.y1;
	const t = Math.min(
		1,
		Math.max(
			0,
			(((pixel.x + 0.5 - x) / width - def.x1) * dx + ((pixel.y + 0.5 - y) / height - def.y1) * dy) /
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
	const from = hexToRgbChannels(a.color)!,
		to = hexToRgbChannels(b.color)!;
	const alpha = (a.opacity ?? 1) + ((b.opacity ?? 1) - (a.opacity ?? 1)) * ratio;
	const colour = (['r', 'g', 'b'] as const).map((k) => from[k] + (to[k] - from[k]) * ratio);
	return { alpha, colour };
}

it('solves native alpha exactly from opaque backdrops and keeps the colour lean small', () => {
	let worstAlpha = 0,
		worstComposite = 0;
	for (const sample of reference.cases) {
		for (const pixel of sample.samples as Pixel[]) {
			const { alpha, colour } = modelAt(sample, pixel);
			for (const [channel, value] of colour.entries()) {
				const black = pixel.onBlack[channel]!,
					white = pixel.onWhite[channel]!;
				worstAlpha = Math.max(worstAlpha, Math.abs((1 - (white - black) / 255 - alpha) * 255));
				worstComposite = Math.max(
					worstComposite,
					Math.abs(value * alpha - black),
					Math.abs(value * alpha + 255 * (1 - alpha) - white),
				);
			}
		}
	}
	// Alpha is straight linear interpolation to rounding; no export artifact is involved.
	expect(worstAlpha).toBeLessThan(1.5);
	// Straight colour interpolation stays within 3.5 levels of an opaque composite. The
	// 3-level translucent limit is not loosened: the remaining lean is the accepted gap.
	expect(worstComposite).toBeLessThan(3.5);
});
