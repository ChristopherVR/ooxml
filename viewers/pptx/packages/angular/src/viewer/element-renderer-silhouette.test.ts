import { readFileSync } from 'node:fs';
import path from 'node:path';

/**
 * A non-rectangular preset paints its fill as a clip-path, which also clipped
 * its shadow and glow. It now paints them as one shape-following SVG filter:
 * the style map carries the `filter` reference, the renderer injects the
 * markup (`getSilhouetteFilterDef`) and the SVG fill overlay paints the shape.
 */
import { buildHollowHitOutline } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getFillLayer, getSilhouetteFilterDef, getSubpathFillOverlay } from './element-effect-defs';
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

	it('keeps a gradient fill and its shadow: the fill is a clipped child layer', () => {
		const style = getShapeFillStrokeStyle(gradientTriangle());
		expect(style['clip-path']).toBeUndefined();
		expect(style['background-image']).toBeUndefined();
		expect(style['filter']).toContain('url(#shape-fx-tri-g)');
		const layer = getFillLayer(gradientTriangle());
		expect(layer?.cssText).toContain('linear-gradient');
		expect(layer?.cssText).toContain('clip-path');
	});

	it('re-enables the text of a text-bearing silhouette shape', () => {
		const text = { ...triangle(), text: 'hello' } as unknown as PptxElement;
		expect(buildHollowHitOutline(text)?.textCss).toContain('pointer-events:auto');
		const html = readFileSync(
			path.join(__dirname, 'element-renderer-shape.component.html'),
			'utf8',
		);
		expect(html).toContain('[css]="hit.textCss"');
	});
});
