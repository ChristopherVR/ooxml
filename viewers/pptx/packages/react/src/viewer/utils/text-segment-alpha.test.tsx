import { buildParagraphs } from 'ooxml-ui/pptx';
import type { PptxElement, TextSegment, TextStyle } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { expect, it } from 'vitest';

import { renderParagraphRun } from './text-segment-render';

it('keeps the theme colour alpha when React resolves a text run colour', () => {
	const segment: TextSegment = {
		text: 'L',
		style: { color: '#FFFFFF', colorRef: { scheme: 'bg1', alpha: 0.3 } },
	};
	const element = {
		id: 'text',
		type: 'text',
		x: 0,
		y: 0,
		width: 80,
		height: 100,
		text: 'L',
		textSegments: [segment],
	} as PptxElement & { textStyle?: TextStyle };
	const run = buildParagraphs(element)[0].runs[0];
	const html = renderToStaticMarkup(
		renderParagraphRun(run, segment, { element, fallbackColor: '#000000' }),
	);
	expect(html).toContain('color:color-mix(in srgb, #FFFFFF 30%, transparent)');
});
