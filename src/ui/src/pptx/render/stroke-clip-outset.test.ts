import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { resolveShapeGeometry } from './shape-geometry-cascade';
import { generateGeometryMorphAnimation } from './morph-geometry-keyframes';
import { outsetClipPathForStroke } from './stroke-clip-outset';
import { buildStrokeOutline, suppressesCssBorder } from './stroke-outline';

/** Whether the point lies in the union of the closed polygons of a flattened `path('...')`. */
function nonZeroContains(clip: string, x: number, y: number): boolean {
	const d = /path\('(.*)'\)/su.exec(clip)?.[1] ?? '';
	let winding = 0;
	for (const sub of d.split('M').filter(Boolean)) {
		const pts = sub.replace(/[LZ]/gu, ' ').trim().split(/\s+/u).map(Number);
		const poly: Array<[number, number]> = [];
		for (let i = 0; i + 1 < pts.length; i += 2) {
			poly.push([pts[i] as number, pts[i + 1] as number]);
		}
		for (let i = 0; i < poly.length; i++) {
			const [x1, y1] = poly[i] as [number, number];
			const [x2, y2] = poly[(i + 1) % poly.length] as [number, number];
			if (y1 <= y ? y2 > y : y2 <= y) {
				const side = (x2 - x1) * (y - y1) - (x - x1) * (y2 - y1);
				if (y1 <= y ? side > 0 : side < 0) {
					winding += y2 > y1 ? 1 : -1;
				}
			}
		}
	}
	return winding !== 0;
}

describe('outsetClipPathForStroke', () => {
	const triangle = "path('M 0 0 L 20 0 L 10 20 Z')";

	it('keeps the outer half of a centred outline inside the clip', () => {
		const widened = outsetClipPathForStroke(triangle, 4, 20, 20);
		expect(nonZeroContains(widened, 10, 5)).toBe(true);
		// 1px above the top edge is inside the outer half of a 4px line.
		expect(nonZeroContains(widened, 10, -1)).toBe(true);
		expect(nonZeroContains(triangle, 10, -1)).toBe(false);
	});

	it('does not reach past half the line width', () => {
		const widened = outsetClipPathForStroke(triangle, 4, 20, 20);
		expect(nonZeroContains(widened, 10, -3)).toBe(false);
	});

	describe('corners', () => {
		const square = "path('M 0 0 L 10 0 L 10 10 L 0 10 Z')";
		const miter = { join: 'miter', miterLimit: 4 } as const;

		it('fills the outer corner of a mitred outline (issue #144)', () => {
			const widened = outsetClipPathForStroke(square, 4, 10, 10, miter);
			// Outside both segment bands, inside the mitre: the notch the bands alone left.
			for (const [x, y] of [
				[10.5, -0.5],
				[11.9, -1.9],
				[-1.9, 11.9],
				[-1.5, -1.5],
			] as const) {
				expect(nonZeroContains(widened, x, y)).toBe(true);
			}
			expect(nonZeroContains(widened, 12.1, -2.1)).toBe(false);
		});

		it('bevels a corner the mitre limit cuts, and a bevel join', () => {
			for (const join of [
				{ join: 'bevel', miterLimit: 4 },
				{ join: 'miter', miterLimit: 1 },
			] as const) {
				const widened = outsetClipPathForStroke(square, 4, 10, 10, join);
				expect(nonZeroContains(widened, -0.9, -0.9)).toBe(true);
				expect(nonZeroContains(widened, -1.5, -1.5)).toBe(false);
			}
		});

		it('rounds the corner of a round join', () => {
			const widened = outsetClipPathForStroke(square, 4, 10, 10, { join: 'round', miterLimit: 4 });
			expect(nonZeroContains(widened, -1.3, -1.3)).toBe(true);
			expect(nonZeroContains(widened, -1.6, -1.6)).toBe(false);
		});

		it('adds nothing past the bands on the inside of a turn', () => {
			// A concave notch: the inner corner at (5, 5) must not grow a mitre into the gap.
			const notch = "path('M 0 0 L 5 0 L 5 5 L 10 5 L 10 10 L 0 10 Z')";
			const widened = outsetClipPathForStroke(notch, 2, 10, 10, miter);
			expect(nonZeroContains(widened, 7, 2)).toBe(false);
			expect(nonZeroContains(widened, 5.5, 4.5)).toBe(true);
		});
	});

	it('returns the clip unchanged for no stroke or unsupported data', () => {
		expect(outsetClipPathForStroke(triangle, 0, 20, 20)).toBe(triangle);
		expect(outsetClipPathForStroke("path('M 0 0 l 5 5')", 2, 20, 20)).toBe("path('M 0 0 l 5 5')");
	});
});

describe('resolveShapeGeometry stroke outset', () => {
	const marker = (shapeStyle: object): PptxElement =>
		({
			id: 'm',
			type: 'shape',
			shapeType: 'flowChartOffpageConnector',
			x: 0,
			y: 0,
			width: 15,
			height: 26,
			shapeStyle: { strokeColor: '#ffffff', ...shapeStyle },
		}) as unknown as PptxElement;

	it('widens the clip of a solid centred outline', () => {
		const solid = resolveShapeGeometry(marker({ strokeWidth: 2 }));
		const none = resolveShapeGeometry(marker({ strokeWidth: 0 }));
		expect(solid.kind).toBe('clipPath');
		expect(none.kind).toBe('clipPath');
		expect((solid as { clipPath: string }).clipPath.length).toBeGreaterThan(
			(none as { clipPath: string }).clipPath.length,
		);
	});

	it('leaves dashed and inset outlines on the exact clip', () => {
		const exact = (resolveShapeGeometry(marker({ strokeWidth: 0 })) as { clipPath: string })
			.clipPath;
		for (const extra of [{ strokeDash: 'dash' }, { lineAlignment: 'in' }]) {
			const geometry = resolveShapeGeometry(marker({ strokeWidth: 2, ...extra }));
			expect((geometry as { clipPath: string }).clipPath).toBe(exact);
		}
	});
});

describe('geometry morph clip widening', () => {
	const shape = (id: string, shapeType: string, strokeWidth: number): PptxElement =>
		({
			id,
			type: 'shape',
			shapeType,
			x: 0,
			y: 0,
			width: 60,
			height: 40,
			shapeStyle: { strokeColor: '#000000', strokeWidth, lineJoin: 'miter' },
		}) as unknown as PptxElement;

	it('widens every baked stop so the outline is not halved mid-tween', () => {
		const animation = generateGeometryMorphAnimation(
			{ fromElement: shape('a', 'rect', 4), toElement: shape('b', 'triangle', 4) },
			1000,
			0,
		);
		const stops = [...(animation?.keyframes ?? '').matchAll(/clip-path: (path\('[^']*'\))/gu)].map(
			(match) => match[1] as string,
		);
		expect(stops.length).toBeGreaterThan(2);
		for (const stop of stops) {
			// The outline polygon plus one band and one join piece per vertex.
			expect(stop.split('M').length - 1).toBeGreaterThan(1);
		}
		// Same command structure at every stop, so CSS can interpolate between them.
		const shapeOf = (clip: string): string => clip.replace(/-?\d*\.?\d+(?:e[-+]?\d+)?/gu, '#');
		expect(new Set(stops.map(shapeOf)).size).toBe(1);
	});

	it('keeps the exact outline when there is no centred stroke', () => {
		const animation = generateGeometryMorphAnimation(
			{ fromElement: shape('a', 'rect', 0), toElement: shape('b', 'triangle', 0) },
			1000,
			0,
		);
		for (const match of (animation?.keyframes ?? '').matchAll(/clip-path: path\('([^']*)'\)/gu)) {
			expect((match[1] as string).split('M').length - 1).toBe(1);
		}
	});
});

describe('inset outlines on clipped shapes', () => {
	const shape = (shapeType: string): PptxElement =>
		({
			id: shapeType,
			type: 'shape',
			shapeType,
			x: 0,
			y: 0,
			width: 80,
			height: 40,
			shapeStyle: { strokeColor: '#ff0000', strokeWidth: 3, lineAlignment: 'in' },
		}) as unknown as PptxElement;

	it('strokes the real outline at twice the width so the exact clip keeps the inner line', () => {
		const outline = buildStrokeOutline(shape('chevron'));
		expect(outline?.strokeWidth).toBe(6);
		expect(suppressesCssBorder(shape('chevron'))).toBe(true);
		// The clip stays exact: widening it would show the outer half too.
		expect((resolveShapeGeometry(shape('chevron')) as { clipPath: string }).clipPath).not.toMatch(
			/Z M/u,
		);
	});

	it('keeps the CSS border for boxes the border already follows', () => {
		for (const type of ['rect', 'roundRect', 'ellipse']) {
			expect(buildStrokeOutline(shape(type))).toBeUndefined();
		}
	});
});
