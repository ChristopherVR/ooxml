import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { ShapeEffectOverlay } from './ShapeEffectOverlay';

/**
 * A resizing shape's Morph re-paths its outline per frame (`morph-shape-carry`
 * in ooxml-ui), through a descendant rule on `[data-pptx-outline] path`. The
 * outline overlay has to carry that marker in every binding or the outline
 * scales with the box instead.
 */
const card = {
	id: 'ppt/slides/slide2.xml-shape-0',
	type: 'shape',
	x: 0,
	y: 0,
	width: 400,
	height: 160,
	shapeType: 'roundRect',
	shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
} as unknown as PptxElement;

describe('shapeEffectOverlay morph outline marker', () => {
	it('stamps data-pptx-outline on the outline overlay that holds the path', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={card} />);
		expect(html).toMatch(/<svg[^>]*data-pptx-outline=""[^>]*>(?:(?!<\/svg>).)*<path/su);
	});
});
