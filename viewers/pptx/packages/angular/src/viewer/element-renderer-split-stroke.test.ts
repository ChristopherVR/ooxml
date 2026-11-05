/**
 * Split-layout presets (an unstroked FILL path plus a separate stroke path: `arc`, `can`,
 * brackets ...). PowerPoint strokes only the stroked path, never the radii of an arc's wedge.
 * Like the other `element-renderer-*.test.ts` files this exercises the accessor the template
 * binds to (`getStrokeOutline`).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getStrokeOutline } from './element-effect-defs';

const arc = {
	id: 'ppt/slides/slide1.xml-shape-1',
	type: 'shape',
	x: 0,
	y: 0,
	width: 200,
	height: 120,
	shapeType: 'arc',
	shapeAdjustments: { adj1: 16200000, adj2: 0 },
	shapeStyle: { fillColor: '#336699', strokeWidth: 4, strokeColor: '#FF0000' },
} as unknown as PptxElement;

describe('angular split-layout preset stroke', () => {
	it('strokes an arc along its curve only, never the wedge radii', () => {
		const outline = getStrokeOutline(arc)!;
		expect(outline.d).toContain('A ');
		expect(outline.d).not.toContain('L ');
	});
});
