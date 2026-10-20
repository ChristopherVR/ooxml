// @vitest-environment happy-dom
import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { renderStrokeOutline } from '../dom/elements/shape-filter-defs';
import { buildCustomPathMarkers, isArrowedCustomPath } from './custom-path-markers';
import { resolveShapeGeometry } from './shape-geometry-cascade';
import { buildStrokeOutline } from './stroke-outline';

function freeform(): PptxElement {
	return {
		id: 'freeform',
		type: 'shape',
		x: 10,
		y: 20,
		width: 200,
		height: 100,
		pathData: 'M 0 1000 L 500 500 Q 1500 0 2000 0',
		pathWidth: 2000,
		pathHeight: 1000,
		shapeStyle: {
			fillMode: 'none',
			strokeWidth: 3,
			strokeColor: '#1F4E79',
			connectorEndArrow: 'triangle',
		},
	};
}

describe('open freeform arrowheads', () => {
	it('paints the scaled path with an endpoint marker and no clipping box', () => {
		const element = freeform();
		expect(resolveShapeGeometry(element).kind).toBe('strokeOnly');
		const outline = buildStrokeOutline(element)!;
		expect(outline.d).toContain('200');
		expect(outline.markers?.markup).toContain('markerWidth="3"');
		expect(outline.markers?.markup).toContain('refX="10"');
		const svg = renderStrokeOutline(document, element)!;
		const path = svg.querySelector(':scope > path')!;
		expect(path.getAttribute('marker-end')).toBe(outline.markers?.end);
		expect(path.getAttribute('stroke-width')).toBe('3');
		expect(svg.querySelectorAll('marker')).toHaveLength(1);
	});
	it('preserves both independently sized ends and stroke opacity', () => {
		const element = freeform();
		if (element.type !== 'shape') throw new Error('expected a shape');
		Object.assign(element.shapeStyle!, {
			connectorStartArrow: 'arrow',
			connectorStartArrowWidth: 'sm',
			connectorEndArrowLength: 'lg',
			strokeOpacity: 0.5,
		});
		const markers = buildStrokeOutline(element)!.markers!;
		expect(markers.start).toBeTruthy();
		expect(markers.markup).toContain('markerWidth="5"');
		expect(markers.markup).toContain('markerHeight="2"');
		expect(markers.markup).toContain('fill="none"');
		expect(markers.markup).toContain('0.5');
	});
	it('leaves closed, filled, and arrowless geometry on its existing path', () => {
		const element = freeform();
		if (element.type !== 'shape') throw new Error('expected a shape');
		element.pathData += ' Z';
		expect(isArrowedCustomPath(element)).toBeFalsy();
		element.pathData = 'M 0 0 L 2000 1000';
		element.shapeStyle = { fillMode: 'solid', fillColor: '#fff', connectorEndArrow: 'triangle' };
		expect(isArrowedCustomPath(element)).toBeFalsy();
		element.shapeStyle = { fillMode: 'none', connectorEndArrow: 'none' };
		expect(buildCustomPathMarkers(element, '#000')).toBeUndefined();
	});
	it('escapes document-derived SVG attributes', () => {
		const markers = buildCustomPathMarkers(freeform(), 'red" onload="bad')!;
		expect(markers.markup).not.toContain('" onload="');
		expect(markers.markup).toContain('&quot;');
	});
});
