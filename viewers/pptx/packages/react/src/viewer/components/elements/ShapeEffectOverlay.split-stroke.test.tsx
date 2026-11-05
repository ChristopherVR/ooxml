import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { renderVectorShape } from '../../utils/vector-shape-renderer';
import { ShapeEffectOverlay } from './ShapeEffectOverlay';

/**
 * Split-layout presets (an unstroked FILL path plus a separate stroke path: `arc`, `can`,
 * brackets ...). PowerPoint strokes only the stroked path, never the radii of an arc's wedge
 * or the chord of a bracket, and fills the wedge when the shape has a fill.
 */
function shape(shapeType: string, extra: Record<string, unknown> = {}): PptxElement {
	return {
		id: 'ppt/slides/slide1.xml-shape-1',
		type: 'shape',
		x: 0,
		y: 0,
		width: 200,
		height: 120,
		shapeType,
		shapeAdjustments: { adj1: 16200000, adj2: 0 },
		shapeStyle: { fillColor: '#336699', strokeWidth: 4, strokeColor: '#FF0000' },
		...extra,
	} as unknown as PptxElement;
}

const strokePaths = (html: string) =>
	[...html.matchAll(/<path d="([^"]*)"[^>]*stroke="#FF0000"/g)].map((m) => m[1]);

describe('split-layout preset stroke (react)', () => {
	it('strokes an arc along its curve only, never the wedge radii', () => {
		const paths = strokePaths(renderToStaticMarkup(<ShapeEffectOverlay element={shape('arc')} />));
		expect(paths.length).toBeGreaterThan(0);
		for (const d of paths) {
			expect(d).toContain('A ');
			expect(d).not.toContain('L ');
		}
	});

	it('paints the wedge of a filled arc as a sub-path fill', () => {
		const html = renderToStaticMarkup(<ShapeEffectOverlay element={shape('arc')} />);
		expect(html).toContain('pptx-react-subpath-fill');
	});

	it('does not hand-draw a can as ellipses and a rect (the shared overlays paint it)', () => {
		expect(renderVectorShape(shape('can'), true, '#336699', 4, '#FF0000')).toBeNull();
	});
});
