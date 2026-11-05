import { passesClicksThroughBox } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { getShapeVisualStyle } from '../../utils/shape-visual-style';
import { shapeParams } from './element-shape-params';
import { ShapeEffectOverlay } from './ShapeEffectOverlay';

/**
 * A non-rectangular preset paints its fill as a clip-path, which also clipped
 * its shadow and glow. It now paints them as one shape-following SVG filter.
 */
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

describe('shape-following shadow (react)', () => {
	it('does not clip the container and puts the shadow in a filter', () => {
		const element = triangle();
		const { hf, fc, sw, sc } = shapeParams(element);
		const style = getShapeVisualStyle(element, hf, fc, sw, sc);
		expect(style.clipPath).toBeUndefined();
		expect(style.boxShadow).toBeUndefined();
		expect(style.filter).toContain('url(#shape-fx-tri-1)');
		expect(style.backgroundColor).toBe('transparent');
	});

	it('injects the filter and paints the fill as an SVG path', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={triangle()} />);
		expect(html).toContain('id="shape-fx-tri-1"');
		expect(html).toContain('pptx-react-subpath-fill');
		expect(html).toContain('fill="#2070C0"');
	});

	it('hit-tests the geometry, not the box corners', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={triangle()} />);
		// The container is pointer-events:none (ElementRenderer, via
		// passesClicksThroughBox); this transparent path opts the triangle back in.
		expect(passesClicksThroughBox(triangle())).toBe(true);
		expect(html).toMatch(/<path[^>]*fill="transparent"[^>]*pointer-events:fill/u);
	});
});
