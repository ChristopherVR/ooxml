import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { renderStrokeOutline } from './shape-filter-defs';

/**
 * Split-layout presets (an unstroked FILL path plus a separate stroke path: `arc`, `can`,
 * brackets ...). PowerPoint strokes only the stroked path, never the radii of an arc's wedge.
 */
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

describe('vanilla split-layout preset stroke', () => {
	it('strokes an arc along its curve only, never the wedge radii', () => {
		const svg = renderStrokeOutline(document, arc);
		const paths = [...(svg?.querySelectorAll('path') ?? [])].map((p) => p.getAttribute('d') ?? '');
		expect(paths.length).toBeGreaterThan(0);
		for (const d of paths) {
			expect(d).toContain('A ');
			expect(d).not.toContain('L ');
		}
	});
});
