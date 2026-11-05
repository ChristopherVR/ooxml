import { mount } from '@vue/test-utils';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import ShapeEffectOverlay from './ShapeEffectOverlay.vue';

/**
 * Split-layout presets (an unstroked FILL path plus a separate stroke path: `arc`, `can`,
 * brackets ...). PowerPoint strokes only the stroked path, never the radii of an arc's wedge,
 * and fills the wedge when the shape has a fill.
 */
function arc(): PptxElement {
	return {
		type: 'shape',
		id: 'sp1',
		x: 0,
		y: 0,
		width: 200,
		height: 120,
		shapeType: 'arc',
		shapeAdjustments: { adj1: 16200000, adj2: 0 },
		shapeStyle: { fillColor: '#336699', strokeWidth: 4, strokeColor: '#FF0000' },
	} as unknown as PptxElement;
}

describe('split-layout preset stroke (vue)', () => {
	it('strokes an arc along its curve only, never the wedge radii', () => {
		const html = mount(ShapeEffectOverlay, { props: { element: arc() } }).html();
		const paths = [...html.matchAll(/<path d="([^"]*)"[^>]*stroke="#FF0000"/gi)].map((m) => m[1]!);
		expect(paths.length).toBeGreaterThan(0);
		for (const d of paths) {
			expect(d).toContain('A ');
			expect(d).not.toContain('L ');
		}
	});
});
