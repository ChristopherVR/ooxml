/**
 * A non-rectangular preset paints its fill as a clip-path, which also clipped
 * its shadow and glow. It now paints them as one shape-following SVG filter:
 * the style map carries the `filter` reference, the renderer injects the
 * markup (`getSilhouetteFilterDef`) and the SVG fill overlay paints the shape.
 */
import { buildHollowHitOutline } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getSilhouetteFilterDef, getSubpathFillOverlay } from './element-effect-defs';
import { getShapeFillStrokeStyle } from './element-style';

function triangle(): PptxElement {
	return {
		id: 'tri-1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 160,
		height: 130,
		shapeType: 'triangle',
		shapeStyle: {
			fillColor: '#2070C0',
			fillMode: 'solid',
			shadowColor: '#000000',
			shadowOpacity: 0.5,
			shadowBlur: 10,
			shadowAngle: 45,
			shadowDistance: 8,
			outerShadowXml: {},
		},
	} as unknown as PptxElement;
}

describe('shape-following shadow (angular)', () => {
	it('does not clip the container and keeps the filter reference', () => {
		const style = getShapeFillStrokeStyle(triangle());
		expect(style['clip-path']).toBeUndefined();
		expect(style['box-shadow']).toBeUndefined();
		expect(style['filter']).toContain('url(#shape-fx-tri-1)');
		expect(style['background-color']).toBe('transparent');
	});

	it('supplies the filter markup and the SVG fill', () => {
		const def = getSilhouetteFilterDef(triangle());
		expect(def?.id).toBe('shape-fx-tri-1');
		expect(def?.filterMarkup).toContain('<filter id="shape-fx-tri-1"');
		const overlay = getSubpathFillOverlay(triangle());
		expect(overlay?.paints.some((p) => p.fill === '#2070C0')).toBeTruthy();
	});

	it('hit-tests the geometry, not the box corners', () => {
		expect(getShapeFillStrokeStyle(triangle())['pointer-events']).toBe('none');
		const hit = buildHollowHitOutline(triangle());
		expect(hit?.pointerEvents).toBe('fill');
		expect(hit?.fill).toBe('transparent');
	});
});
