import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { renderVectorShape } from './vector-shape-renderer';

/**
 * PowerPoint starts a dashed closed outline at the last vertex of the path and
 * runs the closing edge first (see `dash-path-start`); the freeform's own SVG
 * must start there too, or it paints its dashes out of phase with every other
 * binding's outline overlay.
 */
describe('renderVectorShape dashed freeform start', () => {
	const freeform = (strokeDash: string | undefined): PptxElement =>
		({
			id: 'ppt/slides/slide2.xml-shape-0',
			type: 'shape',
			shapeType: 'custom',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			pathData: 'M 0 0 L 100 0 L 100 50 L 0 50 Z',
			pathWidth: 100,
			pathHeight: 50,
			shapeStyle: { strokeColor: '#FFFFFF', strokeWidth: 8, strokeDash },
		}) as unknown as PptxElement;

	it('starts the dashed stroke at the last vertex', () => {
		const html = renderToStaticMarkup(renderVectorShape(freeform('dash'), false, '', 8, '#FFFFFF'));
		expect(html).toContain('d="M 0 50 L 0 0 L 100 0 L 100 50 L 0 50 Z"');
	});

	it('keeps a solid stroke in path order', () => {
		const html = renderToStaticMarkup(
			renderVectorShape(freeform(undefined), false, '', 8, '#FFFFFF'),
		);
		expect(html).toContain('d="M 0 0 L 100 0 L 100 50 L 0 50 Z"');
	});
});
