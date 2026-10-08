import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { parseVsdx } from '../parser';
import { visioFilledArrow } from '../filled-arrow';
import { visioSvgStrokeStyle } from './svg-stroke';
import { visioLineDashLengths } from './line-dash';

const directory = process.env.VISIO_NATIVE_SVG_ZERO_STROKE_DIR;
interface NativeStroke {
	pageId: string;
	shapeId: string;
	weight: number;
	cap: number;
	arrow: boolean;
	svgStyle: string;
	markers: { reference: number; transform: string }[];
}
it.skipIf(!directory)(
	'matches native zero-weight SVG strokes, dots and filled arrow setbacks',
	async () => {
		const evidence = JSON.parse(
			(await readFile(join(directory!, 'evidence.json'), 'utf8')).trim(),
		) as {
			lines: NativeStroke[];
		};
		const document = await parseVsdx(
			new Uint8Array(await readFile(join(directory!, 'zero-stroke.vsdx'))),
		);
		expect(evidence.lines).toHaveLength(12);
		for (const entry of evidence.lines) {
			const shape = document.pages
				.find((page) => page.id === entry.pageId)!
				.shapes.find((candidate) => candidate.id === entry.shapeId)!;
			const rawWidth = shape.style.lineWidth;
			const output = visioSvgStrokeStyle(shape.style);
			const nativeWidth = Number(/stroke-width:([^;}]+)/u.exec(entry.svgStyle)![1]);
			expect(rawWidth).toBeCloseTo(entry.weight / 72, 12);
			expect(output.lineWidth * 72).toBeCloseTo(nativeWidth, 12);
			expect(shape.style.lineWidth).toBe(rawWidth);
			expect(shape.style.lineCap).toBe(['round', 'butt', 'square'][entry.cap]);
			const nativeDash = /stroke-dasharray:([^;}]+)/u
				.exec(entry.svgStyle)?.[1]
				?.split(',')
				.map(Number);
			const dash = visioLineDashLengths(output);
			if (nativeDash) {
				expect(dash).toHaveLength(nativeDash.length);
				for (let index = 0; index < nativeDash.length; index++)
					expect(dash![index]! * 72).toBeCloseTo(nativeDash[index]!, 9);
			} else expect(dash).toBeUndefined();
			if (entry.arrow) {
				const arrow = visioFilledArrow(
					output.endArrow,
					output.endArrowSize ?? 2,
					output.lineWidth,
				)!;
				expect(entry.markers).toHaveLength(2);
				expect(arrow.beginSetback * 72).toBeCloseTo(entry.markers[0]!.reference * nativeWidth, 8);
				expect(arrow.setback * 72).toBeCloseTo(-entry.markers[1]!.reference * nativeWidth, 8);
				if (entry.weight === 0) {
					const explicit = evidence.lines.find((line) => line.arrow && line.weight === 0.75)!;
					expect(entry.markers).toEqual(explicit.markers);
				}
			}
		}
	},
	120_000,
);
