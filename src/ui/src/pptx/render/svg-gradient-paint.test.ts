import type { ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildSvgGradientDef,
	svgGradientFillRef,
	svgGradientId,
	svgGradientMarkup,
} from './svg-gradient-paint';

const twoStops = [
	{ color: '#D2F8F4', position: 0 },
	{ color: '#FFFFFF', position: 100 },
];

function style(overrides: ShapeStyle = {}): ShapeStyle {
	return { fillMode: 'gradient', fillGradientStops: twoStops, ...overrides } as ShapeStyle;
}

describe('svgGradientId', () => {
	it('sanitises an element id into a usable fragment', () => {
		// Raw element ids carry `/` and `.`, which cannot appear in `url(#…)`.
		expect(svgGradientId('ppt/slides/slide2.xml-shape-4')).toBe(
			'pptx-grad-ppt_slides_slide2_xml-shape-4',
		);
	});
});

describe('buildSvgGradientDef', () => {
	it('returns undefined for a non-gradient fill', () => {
		expect(buildSvgGradientDef({ fillMode: 'solid', fillColor: '#ff0000' }, 'e1')).toBeUndefined();
	});

	it('returns undefined when there are no structured stops', () => {
		expect(
			buildSvgGradientDef(
				{ fillMode: 'gradient', fillGradient: 'linear-gradient(red, blue)' },
				'e1',
			),
		).toBeUndefined();
	});

	it('runs a 0-degree linear gradient left to right', () => {
		// OOXML `a:lin/@ang="0"` runs along +x; SVG shares that axis convention,
		// so no CSS-style quarter turn is applied.
		const def = buildSvgGradientDef(style({ fillGradientAngle: 0 }), 'e1');
		expect(def).toMatchObject({ kind: 'linear', x1: 0, y1: 0.5, x2: 1, y2: 0.5 });
	});

	it('runs a 90-degree linear gradient top to bottom', () => {
		const def = buildSvgGradientDef(style({ fillGradientAngle: 90 }), 'e1');
		expect(def).toMatchObject({ kind: 'linear', x1: 0.5, y1: 0, x2: 0.5, y2: 1 });
	});

	it('spans corner to corner at 45 degrees', () => {
		const def = buildSvgGradientDef(style({ fillGradientAngle: 45 }), 'e1');
		expect(def).toMatchObject({ kind: 'linear', x1: 0, y1: 0, x2: 1, y2: 1 });
	});

	it('defaults to top-to-bottom when no angle is recorded', () => {
		expect(buildSvgGradientDef(style(), 'e1')).toMatchObject({ x1: 0.5, y1: 0, x2: 0.5, y2: 1 });
	});

	it('converts stop positions to 0-1 offsets and keeps alpha', () => {
		const def = buildSvgGradientDef(
			style({
				fillGradientStops: [
					{ color: '#000000', position: 0 },
					{ color: '#808080', position: 50 },
					{ color: '#FFFFFF', position: 100, opacity: 0 },
				],
			}),
			'e1',
		);
		expect(def?.stops).toStrictEqual([
			{ offset: 0, color: '#000000' },
			{ offset: 0.5, color: '#808080' },
			{ offset: 1, color: '#FFFFFF', opacity: 0 },
		]);
	});

	it('re-samples a two-stop 0/100 pair onto the Office sigma curve', () => {
		// PowerPoint paints white to black 186 (not 128) at the midpoint.
		const def = buildSvgGradientDef(style(), 'e1');
		const stops = def!.stops;
		expect(stops.length).toBeGreaterThan(20);
		const mid = stops.find((stop) => Math.abs(stop.offset - 0.5) < 0.01)!;
		expect(Number.parseInt(mid.color.slice(1, 3), 16)).toBeGreaterThan(170);
	});

	it('centres a radial gradient on its fillToRect', () => {
		// `<a:fillToRect r="100000" b="100000"/>` collapses to the top-left corner.
		const def = buildSvgGradientDef(
			style({
				fillGradientType: 'radial',
				fillGradientPathType: 'circle',
				fillGradientFillToRect: { l: 0, t: 0, r: 1, b: 1 },
			}),
			'e1',
		);
		expect(def).toMatchObject({ kind: 'radial', cx: 0, cy: 0 });
		// Farthest corner of the unit box from (0, 0).
		expect((def as { r: number }).r).toBeCloseTo(Math.SQRT2, 4);
	});

	it('centres a radial gradient with no fillToRect on the shape', () => {
		const def = buildSvgGradientDef(style({ fillGradientType: 'radial' }), 'e1');
		expect(def).toMatchObject({ kind: 'radial', cx: 0.5, cy: 0.5 });
	});

	it('builds a rectPath pattern def (not an ellipse radialGradient) for path type "rect"', () => {
		const def = buildSvgGradientDef(
			style({ fillGradientType: 'radial', fillGradientPathType: 'rect' }),
			'e1',
		);
		expect(def).toMatchObject({ kind: 'rectPath', id: 'pptx-grad-e1' });
		// A bare data URI (valid for an SVG `<image href="...">` attribute), NOT
		// the CSS `url("...")` wrapper `buildRectPathGradientImage` returns.
		expect((def as { href: string }).href).toMatch(/^data:image\/svg\+xml,/u);
		// The nested-rectangle band SVG itself, embedded in the data URI.
		expect(decodeURIComponent((def as { href: string }).href)).toContain('<polygon ');
	});

	it('keeps path type "shape" elliptical (only "rect" gets the band field)', () => {
		const def = buildSvgGradientDef(
			style({ fillGradientType: 'radial', fillGradientPathType: 'shape' }),
			'e1',
		);
		expect(def).toMatchObject({ kind: 'radial' });
	});
});

describe('svgGradientFillRef / svgGradientMarkup', () => {
	it('references the definition by id', () => {
		const def = buildSvgGradientDef(style(), 'ppt/slides/slide3.xml-shape-1');
		expect(svgGradientFillRef(def!)).toBe('url(#pptx-grad-ppt_slides_slide3_xml-shape-1)');
	});

	it('serialises a linear definition to markup', () => {
		const def = buildSvgGradientDef(style({ fillGradientAngle: 0 }), 'e1');
		const markup = svgGradientMarkup(def!);
		expect(markup).toContain('<linearGradient id="pptx-grad-e1"');
		expect(markup).toContain('<stop offset="0" stop-color="#D2F8F4"/>');
		expect(markup).toContain('<stop offset="1" stop-color="#FFFFFF"/>');
	});

	it('serialises a transparent stop with stop-opacity', () => {
		const def = buildSvgGradientDef(
			style({
				fillGradientStops: [
					{ color: '#000000', position: 0 },
					{ color: '#FFFFFF', position: 100, opacity: 0 },
				],
			}),
			'e1',
		);
		expect(svgGradientMarkup(def!)).toContain('stop-opacity="0"');
	});

	it('serialises a radial definition to markup', () => {
		const def = buildSvgGradientDef(style({ fillGradientType: 'radial' }), 'e1');
		expect(svgGradientMarkup(def!)).toContain('<radialGradient id="pptx-grad-e1"');
	});

	it('serialises a rectPath definition to a stretched <pattern><image> tile', () => {
		const def = buildSvgGradientDef(
			style({ fillGradientType: 'radial', fillGradientPathType: 'rect' }),
			'e1',
		);
		const markup = svgGradientMarkup(def!);
		expect(markup).toContain('<pattern id="pptx-grad-e1" patternUnits="objectBoundingBox"');
		// The tile's own content must be in bounding-box units too, else the 1x1 image
		// is one user unit big and the fill paints nothing (React's blank rect path).
		expect(markup).toContain('patternContentUnits="objectBoundingBox"');
		expect(markup).toContain('width="1" height="1"');
		expect(markup).toContain('preserveAspectRatio="none"');
		expect(markup).not.toContain('<radialGradient');
	});
});

describe('buildSvgGradientDef with the shape size (freeform)', () => {
	const bounds = { width: 400, height: 200 };

	it('paints path="shape" and path="circle" as a true circle to the farthest corner', () => {
		for (const path of ['shape', 'circle'] as const) {
			const def = buildSvgGradientDef(
				style({ fillGradientType: 'radial', fillGradientPathType: path }),
				'e1',
				bounds,
			);
			expect(def?.kind).toBe('rectPath');
			const svg = decodeURIComponent((def as { href: string }).href);
			// radius is half the diagonal in the box's own pixels, not the 0.707 of a stretched box
			expect(svg).toContain('r="223.6068"');
			expect(svg).toContain('viewBox="0 0 400 200"');
		}
	});

	it('keeps an off-centre focus on the circumscribed circle', () => {
		const def = buildSvgGradientDef(
			style({
				fillGradientType: 'radial',
				fillGradientPathType: 'shape',
				fillGradientFillToRect: { l: 0, t: 0, r: 1, b: 1 },
			}),
			'e1',
			bounds,
		);
		const svg = decodeURIComponent((def as { href: string }).href);
		expect(svg).toContain('fx="0" fy="0"');
		expect(svg).toContain('cx="200" cy="100"');
	});

	it('keeps the bounding-box ellipse when the size is unknown', () => {
		const def = buildSvgGradientDef(
			style({ fillGradientType: 'radial', fillGradientPathType: 'shape' }),
			'e1',
		);
		expect(def?.kind).toBe('radial');
	});

	it('honours scaled="0": a true angle on a wide box', () => {
		const def = buildSvgGradientDef(
			style({ fillGradientType: 'linear', fillGradientAngle: 45, fillGradientScaled: false }),
			'e1',
			bounds,
		);
		expect(def?.kind).toBe('rectPath');
		const svg = decodeURIComponent((def as { href: string }).href);
		// 45 degrees through the centre, spanning the box's projection (|cos| w + |sin| h) / 2
		expect(svg).toContain('x1="50"');
		expect(svg).toContain('y1="-50"');
		// the default scaled gradient stays a bounding-box linear gradient
		expect(
			buildSvgGradientDef(
				style({ fillGradientType: 'linear', fillGradientAngle: 45 }),
				'e1',
				bounds,
			)?.kind,
		).toBe('linear');
	});
});
