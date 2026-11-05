import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { renderStrokeOutline } from './shape-filter-defs';

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

const tri = {
	...card,
	shapeStyle: { ...card.shapeStyle, strokeWidth: 12, compoundLine: 'tri' },
} as unknown as PptxElement;

describe('vanilla stroke outline morph marker', () => {
	it('stamps data-pptx-outline on the outline overlay that holds the path', () => {
		const svg = renderStrokeOutline(document, card);
		expect(svg?.hasAttribute('data-pptx-outline')).toBe(true);
		expect(svg?.querySelector('path')).not.toBeNull();
	});

	it('paints each strand of a compound outline as a sibling path, in order', () => {
		// The compound carry addresses strands as `[data-pptx-outline] path:nth-of-type(n)`.
		const svg = renderStrokeOutline(document, tri);
		expect(svg?.querySelectorAll(':scope > path')).toHaveLength(3);
	});
});
