import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { renderVectorShape } from './vector-shape-renderer';

/**
 * A freeform is the one shape React paints as a vector `<svg>` in path units
 * rather than a clipped box. A morph that reshapes the freeform re-paths it per
 * frame through the shared plan's `customGeometryAnimations`, whose rules
 * select `[data-pptx-custom-geometry] path`; without the marker the fill kept
 * its final outline under the morphing clip and its line stayed on screen
 * beside the carried outline.
 */
describe('renderVectorShape custom geometry marker', () => {
	it('marks the freeform svg so a morph can re-path it', () => {
		const element = {
			id: 'ppt/slides/slide2.xml-shape-0',
			type: 'shape',
			shapeType: 'custom',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			pathData: 'M 0 0 L 100 0 L 100 50 Z',
			pathWidth: 100,
			pathHeight: 50,
			shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
		} as unknown as PptxElement;
		const html = renderToStaticMarkup(renderVectorShape(element, true, '#C07020', 8, '#FFFFFF'));

		expect(html).toMatch(/<svg data-pptx-custom-geometry="" viewBox="0 0 100 50"/u);
	});
});
