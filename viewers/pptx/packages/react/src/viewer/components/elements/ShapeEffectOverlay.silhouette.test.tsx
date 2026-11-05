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

function gradientTriangle(): PptxElement {
	return {
		id: 'tri-g',
		type: 'shape',
		x: 0,
		y: 0,
		width: 160,
		height: 130,
		shapeType: 'triangle',
		shapeStyle: {
			fillColor: '#2070C0',
			fillMode: 'gradient',
			fillGradientType: 'linear',
			fillGradientAngle: 90,
			fillGradientStops: [
				{ color: '#FF0000', position: 0 },
				{ color: '#0000FF', position: 100 },
			],
			shadowColor: '#000000',
			shadowOpacity: 0.5,
			shadowBlur: 10,
			shadowAngle: 45,
			shadowDistance: 8,
			outerShadowXml: {},
		},
	} as unknown as PptxElement;
}

function innerShadow(shapeType: string): PptxElement {
	return {
		id: 'in-1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 160,
		height: 130,
		shapeType,
		shapeStyle: {
			fillColor: '#2070C0',
			fillMode: 'solid',
			innerShadowColor: '#000000',
			innerShadowOpacity: 0.7,
			innerShadowBlur: 8,
			innerShadowOffsetX: 8,
			innerShadowOffsetY: 0,
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

	it('keeps a gradient fill and its shadow: the fill is a clipped child layer', () => {
		const element = gradientTriangle();
		const { hf, fc, sw, sc } = shapeParams(element);
		const style = getShapeVisualStyle(element, hf, fc, sw, sc);
		expect(style.clipPath).toBeUndefined();
		expect(style.backgroundImage).toBeUndefined();
		expect(style.filter).toContain('url(#shape-fx-tri-g)');
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={element} />);
		expect(html).toContain('pptx-react-silhouette-fill');
		expect(html).toContain('linear-gradient');
		expect(html).toContain('clip-path');
	});

	it('re-enables the text of a text-bearing silhouette shape', () => {
		const text = { ...triangle(), text: 'hello' } as unknown as PptxElement;
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={text} />);
		expect(html).toContain('<style>');
		expect(html).toContain('span');
		expect(html).toContain('pointer-events:auto');
	});

	it('paints the inner shadow of a clipped shape in the filter, PowerPoint side up', () => {
		const tri = innerShadow('triangle');
		const t = shapeParams(tri);
		const triStyle = getShapeVisualStyle(tri, t.hf, t.fc, t.sw, t.sc);
		expect(triStyle.filter).toContain('url(#shape-fx-in-1)');
		expect(triStyle.boxShadow).toBeUndefined();
		const rect = innerShadow('rect');
		const r = shapeParams(rect);
		expect(getShapeVisualStyle(rect, r.hf, r.fc, r.sw, r.sc).boxShadow).toContain('inset -8px');
	});

	it('marks the silhouette fill path so a morph can re-path it', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={triangle()} />);
		expect(html).toContain('data-pptx-silhouette-fill');
	});
});
