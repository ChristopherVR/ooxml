import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { resolveShapeGeometry } from './shape-geometry-cascade';
import { outsetClipPathForStroke } from './stroke-clip-outset';

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
