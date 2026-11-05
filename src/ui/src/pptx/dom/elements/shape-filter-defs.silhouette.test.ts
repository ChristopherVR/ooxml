// @vitest-environment happy-dom
import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { getShapeFillStrokeStyle } from '../element-styles';
import {
	renderShapeFillLayer,
	renderHollowHitOutline,
	renderShapeFilterDefs,
	renderShapeSubpathFillOverlay,
} from './shape-filter-defs';

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

describe('shape-following shadow (vanilla)', () => {
	it('does not clip the container and puts the shadow in a filter', () => {
		const style = getShapeFillStrokeStyle(triangle());
		expect(style['clipPath']).toBeUndefined();
		expect(style['boxShadow']).toBeUndefined();
		expect(String(style['filter'])).toContain('url(#shape-fx-tri-1)');
		expect(style['backgroundColor']).toBe('transparent');
	});

	it('injects the filter and paints the fill as an SVG path', () => {
		const defs = renderShapeFilterDefs(document, triangle());
		expect(defs?.innerHTML).toContain('id="shape-fx-tri-1"');
		const svg = renderShapeSubpathFillOverlay(document, triangle());
		expect(
			[...(svg?.querySelectorAll('path') ?? [])].some((p) => p.getAttribute('fill') === '#2070C0'),
		).toBe(true);
	});

	it('hit-tests the geometry, not the box corners', () => {
		expect(getShapeFillStrokeStyle(triangle())['pointerEvents']).toBe('none');
		const svg = renderHollowHitOutline(document, triangle());
		const path = svg?.querySelector('path');
		expect(path?.getAttribute('fill')).toBe('transparent');
		expect(path?.getAttribute('style')).toContain('pointer-events:fill');
	});

	it('keeps a gradient fill and its shadow: the fill is a clipped child layer', () => {
		const style = getShapeFillStrokeStyle(gradientTriangle());
		expect(style['clipPath']).toBeUndefined();
		expect(style['backgroundImage']).toBeUndefined();
		expect(String(style['filter'])).toContain('url(#shape-fx-tri-g)');
		const layer = renderShapeFillLayer(document, gradientTriangle());
		expect(layer?.style.backgroundImage).toContain('linear-gradient');
		expect(layer?.style.clipPath).toBeTruthy();
	});

	it('re-enables the text of a text-bearing silhouette shape', () => {
		const text = { ...triangle(), text: 'hello' } as unknown as PptxElement;
		const svg = renderHollowHitOutline(document, text);
		expect(svg?.querySelector('style')?.textContent).toContain('pointer-events:auto');
	});
});
