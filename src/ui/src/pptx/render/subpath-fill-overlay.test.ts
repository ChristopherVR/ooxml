import type { CustomGeometryPath, PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildSubpathFillOverlay, suppressesCssFill } from './subpath-fill-overlay';

function shape(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'shape',
		id: 'sp1',
		x: 0,
		y: 0,
		width: 200,
		height: 200,
		shapeStyle: { fillColor: '#336699' },
		...overrides,
	} as PptxElement;
}

describe('suppressesCssFill / buildSubpathFillOverlay - preset geometry (bug 2)', () => {
	it('fires for smileyFace: the eyes are open (fill="none") strokes, not filled', () => {
		const el = shape({ shapeType: 'smileyFace' });
		expect(suppressesCssFill(el)).toBeTruthy();
		const overlay = buildSubpathFillOverlay(el)!;
		expect(overlay).toBeDefined();
		// Face (first sub-path) is filled with the shape colour...
		expect(overlay.paints[0].fill).not.toBe('none');
		// ...and at least one sub-path (an eye) opts out of fill entirely.
		expect(overlay.paints.some((p) => p.fill === 'none')).toBeTruthy();
	});

	it('fires for actionButtonForwardNext: the inset bevel well is darkened, not flat', () => {
		const el = shape({ shapeType: 'actionButtonForwardNext' });
		expect(suppressesCssFill(el)).toBeTruthy();
		const overlay = buildSubpathFillOverlay(el)!;
		const fills = new Set(overlay.paints.map((p) => p.fill));
		// The darkened bevel well must resolve to a DIFFERENT colour than the
		// plain face fill, or the shading is invisible (the bug as filed).
		expect(fills.size).toBeGreaterThan(1);
	});

	it('declines an ordinary single-fill preset (rect): the CSS box stays correct and cheaper', () => {
		const el = shape({ shapeType: 'rect' });
		expect(suppressesCssFill(el)).toBeFalsy();
		expect(buildSubpathFillOverlay(el)).toBeUndefined();
	});

	// PowerPoint (COM export) fills an arc's wedge when the shape has a fill and strokes only the
	// curve: the ECMA-376 `arc` is an unstroked wedge fill path plus a fill="none" stroke path, so
	// it is a per-sub-path fill like smileyFace, not an open preset.
	it('paints the wedge of a filled arc and nothing for its fill="none" curve', () => {
		const el = shape({ shapeType: 'arc', shapeAdjustments: { adj1: 16_200_000, adj2: 0 } });
		expect(suppressesCssFill(el)).toBeTruthy();
		const overlay = buildSubpathFillOverlay(el)!;
		expect(overlay.paints[0].fill).not.toBe('none');
		expect(overlay.paints[1].fill).toBe('none');
	});

	it('paints no fill for an unfilled arc', () => {
		const el = shape({
			shapeType: 'arc',
			shapeStyle: { fillMode: 'none' },
			shapeAdjustments: { adj1: 16_200_000, adj2: 0 },
		} as Partial<PptxElement>);
		expect(buildSubpathFillOverlay(el)!.paints.every((p) => p.fill === 'none')).toBeTruthy();
	});

	it('keeps the CSS gradient and veils only the shaded sub-paths (the eyes of a smiley)', () => {
		const el = shape({
			shapeType: 'smileyFace',
			shapeStyle: { fillMode: 'gradient', fillGradientStops: [{ color: '#fff', position: 0 }] },
		} as Partial<PptxElement>);
		expect(suppressesCssFill(el)).toBeFalsy();
		const overlay = buildSubpathFillOverlay(el)!;
		expect(overlay).toBeDefined();
		// face plain, eyes a 20% black veil (darkenLess), mouth nothing
		expect(overlay.paints[0].fill).toBe('none');
		expect(overlay.paints.some((p) => p.fill === 'rgba(0,0,0,0.2)')).toBeTruthy();
	});

	it('veils a can lid white and a cube side black over a gradient or a picture', () => {
		const style = { fillMode: 'image', fillImageUrl: 'x.png' };
		const can = buildSubpathFillOverlay(
			shape({ shapeType: 'can', shapeStyle: style } as Partial<PptxElement>),
		)!;
		expect(can.paints.some((p) => p.fill === 'rgba(255,255,255,0.4)')).toBeTruthy();
		const cube = buildSubpathFillOverlay(
			shape({
				shapeType: 'cube',
				shapeStyle: { ...style, fillMode: 'gradient' },
			} as Partial<PptxElement>),
		)!;
		expect(cube.paints.some((p) => p.fill === 'rgba(0,0,0,0.2)')).toBeTruthy();
		expect(cube.paints.some((p) => p.fill === 'rgba(255,255,255,0.2)')).toBeTruthy();
	});

	it('declines a gradient shape with no shaded sub-path', () => {
		const el = shape({
			shapeType: 'rect',
			shapeStyle: { fillMode: 'gradient', fillGradientStops: [{ color: '#fff', position: 0 }] },
		} as Partial<PptxElement>);
		expect(buildSubpathFillOverlay(el)).toBeUndefined();
	});

	it('emits no fill for every sub-path when the shape itself has no fill', () => {
		const el = shape({ shapeType: 'smileyFace', shapeStyle: { fillMode: 'none' } });
		const overlay = buildSubpathFillOverlay(el)!;
		expect(overlay.paints.every((p) => p.fill === 'none')).toBeTruthy();
	});
});

describe('suppressesCssFill / buildSubpathFillOverlay - custom geometry (bug 3)', () => {
	function custGeomPaths(): CustomGeometryPath[] {
		return [
			{
				width: 100,
				height: 100,
				segments: [
					{ type: 'moveTo', pt: { x: 0, y: 0 } },
					{ type: 'lineTo', pt: { x: 100, y: 0 } },
					{ type: 'lineTo', pt: { x: 100, y: 100 } },
					{ type: 'close' },
				],
			},
			{
				width: 100,
				height: 100,
				fillMode: 'lighten',
				segments: [
					{ type: 'moveTo', pt: { x: 10, y: 10 } },
					{ type: 'lineTo', pt: { x: 40, y: 10 } },
					{ type: 'lineTo', pt: { x: 40, y: 40 } },
					{ type: 'close' },
				],
			},
		];
	}

	function custGeomShape(): PptxElement {
		return shape({
			pathData: 'M 0 0 L 100 0 L 100 100 Z M 10 10 L 40 10 L 40 40 Z',
			pathWidth: 100,
			pathHeight: 100,
			customGeometryPaths: custGeomPaths(),
		} as Partial<PptxElement>);
	}

	it('fires when a structured sub-path carries a non-norm fill mode', () => {
		const el = custGeomShape();
		expect(suppressesCssFill(el)).toBeTruthy();
		const overlay = buildSubpathFillOverlay(el)!;
		expect(overlay.paints).toHaveLength(2);
		expect(overlay.paints[0].fill).not.toBe(overlay.paints[1].fill);
		expect(overlay.viewBoxWidth).toBe(100);
		expect(overlay.viewBoxHeight).toBe(100);
	});

	it('declines custom geometry with no structured sub-paths (aggregate pathData only)', () => {
		const el = shape({
			pathData: 'M 0 0 L 100 0 L 100 100 Z',
			pathWidth: 100,
			pathHeight: 100,
		} as Partial<PptxElement>);
		expect(suppressesCssFill(el)).toBeFalsy();
	});

	it('declines custom geometry whose sub-paths are all plain (norm, stroked)', () => {
		const el = shape({
			pathData: 'M 0 0 L 100 0 L 100 100 Z',
			pathWidth: 100,
			pathHeight: 100,
			customGeometryPaths: [
				{
					width: 100,
					height: 100,
					segments: [
						{ type: 'moveTo', pt: { x: 0, y: 0 } },
						{ type: 'lineTo', pt: { x: 100, y: 0 } },
						{ type: 'close' },
					],
				},
			],
		} as Partial<PptxElement>);
		expect(suppressesCssFill(el)).toBeFalsy();
	});

	it('is unaffected by element type: only shape/image/picture carry custom geometry', () => {
		const el = { ...custGeomShape(), type: 'table' } as unknown as PptxElement;
		expect(suppressesCssFill(el)).toBeFalsy();
	});
});
