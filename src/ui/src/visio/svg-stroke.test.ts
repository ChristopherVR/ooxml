import { expect, it } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { exportPageSvg } from './export-svg';
import { renderPage } from './render-svg';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';

function drawing(lineWidth: number, pattern = 1) {
	const model = structuredClone(demoDocument);
	const page = model.pages[0]!;
	const line = page.shapes.find((shape) => shape.kind === 'connector')!;
	page.shapes = [line];
	Object.assign(line.style, {
		lineWidth,
		linePattern: pattern,
		lineCap: 'round',
		startArrow: 4,
		endArrow: 4,
		...(pattern === 3 ? { lineDash: [0, 5] } : {}),
	});
	return { model, page, line };
}

it('exports zero-weight SVG strokes and marker geometry like explicit 0.75-point strokes', () => {
	const zero = drawing(0);
	const explicit = drawing(0.75 / 72);
	const parse = (model: typeof zero.model) =>
		new DOMParser().parseFromString(exportPageSvg(model).svg, 'image/svg+xml');
	const actual = parse(zero.model),
		reference = parse(explicit.model);
	const geometry = (xml: Document) =>
		[...xml.querySelectorAll('path')].map((path) => ({
			path: path.getAttribute('d'),
			width: path.getAttribute('stroke-width'),
		}));
	expect(geometry(actual)).toEqual(geometry(reference));
	expect(actual.querySelector('[stroke-width]')?.getAttribute('stroke-width')).toBe(
		String(0.75 / 72),
	);
	expect(actual.querySelectorAll('marker')).toHaveLength(2);
	expect(zero.line.style.lineWidth).toBe(0);
	const live = renderPage(zero.model, zero.page);
	expect(live.svg.querySelector('[data-geometry]')?.getAttribute('stroke-width')).toBe('0');
	live.dispose();
});

it.each([0, 0.001 / 72])(
	'uses output width for dashed %s-inch strokes without clamping nonzero weights',
	(width) => {
		const { model, line } = drawing(width, 3);
		const xml = new DOMParser().parseFromString(exportPageSvg(model).svg, 'image/svg+xml');
		const path = xml.querySelector('[stroke-dasharray]')!;
		const effective = width === 0 ? 0.75 / 72 : width;
		expect(path.getAttribute('stroke-width')).toBe(String(effective));
		expect(path.getAttribute('stroke-dasharray')).toBe(`0 ${5 * effective}`);
		expect(line.style.lineWidth).toBe(width);
	},
);

it('keeps no-line geometry and its markers suppressed at zero weight', () => {
	const { model } = drawing(0, 0);
	const xml = new DOMParser().parseFromString(exportPageSvg(model).svg, 'image/svg+xml');
	expect(xml.querySelectorAll('marker')).toHaveLength(0);
	expect(xml.querySelector('[stroke-width]')?.getAttribute('stroke')).toBe('none');
});

const nativeDirectory = process.env.VISIO_NATIVE_SVG_ZERO_STROKE_DIR;
it.skipIf(!nativeDirectory)(
	'exports the twelve actual native stroke fixtures through the shared DOM renderer',
	async () => {
		const evidence = JSON.parse(
			(await readFile(join(nativeDirectory!, 'evidence.json'), 'utf8')).trim(),
		) as {
			lines: {
				pageId: string;
				svgStyle: string;
				arrow: boolean;
				markers: { reference: number }[];
			}[];
		};
		const model = await parseVsdx(
			new Uint8Array(await readFile(join(nativeDirectory!, 'zero-stroke.vsdx'))),
		);
		expect(evidence.lines).toHaveLength(12);
		for (const entry of evidence.lines) {
			const index = model.pages.findIndex((page) => page.id === entry.pageId);
			const rawWidth = model.pages[index]!.shapes[0]!.style.lineWidth;
			const xml = new DOMParser().parseFromString(exportPageSvg(model, index).svg, 'image/svg+xml');
			const path = xml.querySelector('[stroke-width]')!;
			const nativeWidth = Number(/stroke-width:([^;}]+)/u.exec(entry.svgStyle)![1]);
			expect(Number(path.getAttribute('stroke-width')) * 72).toBeCloseTo(nativeWidth, 12);
			const nativeDash = /stroke-dasharray:([^;}]+)/u
				.exec(entry.svgStyle)?.[1]
				?.split(',')
				.map(Number);
			if (nativeDash) {
				const actual = path.getAttribute('stroke-dasharray')!.split(' ').map(Number);
				expect(actual).toHaveLength(nativeDash.length);
				for (let i = 0; i < actual.length; i++)
					expect(actual[i]! * 72).toBeCloseTo(nativeDash[i]!, 9);
			}
			if (entry.arrow) {
				const markers = [...xml.querySelectorAll('marker')];
				expect(markers).toHaveLength(2);
				expect(-Number(markers[0]!.getAttribute('refX')) * 72).toBeCloseTo(
					entry.markers[0]!.reference * nativeWidth,
					8,
				);
				expect(Number(markers[1]!.getAttribute('refX')) * 72).toBeCloseTo(
					entry.markers[1]!.reference * nativeWidth,
					8,
				);
			}
			expect(model.pages[index]!.shapes[0]!.style.lineWidth).toBe(rawWidth);
		}
	},
);
