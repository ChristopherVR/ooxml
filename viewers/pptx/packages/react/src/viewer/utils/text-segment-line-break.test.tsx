import { buildParagraphs } from 'ooxml-ui/pptx';
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { expect, it } from 'vitest';

import { renderTextSegments } from './text-paragraph-render';

it('renders an a:br as a bare <br>, not a span wrapping one', () => {
	const segments: TextSegment[] = [
		{ text: 'A SPECIAL ', style: {} },
		{ text: '\n', style: {}, isLineBreak: true } as TextSegment,
		{ text: 'DAY IN', style: {} },
	];
	const element = {
		id: 'title',
		type: 'text',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		text: 'A SPECIAL \nDAY IN',
		textSegments: segments,
	} as PptxElement;
	// Guard the premise: shared emits the break as its own '\n' run.
	expect(buildParagraphs(element)[0].runs.some((r) => r.text === '\n')).toBe(true);
	const html = renderToStaticMarkup(<>{renderTextSegments(element, '#000000')}</>);
	expect(html).toContain('<br data-seg-idx="1"/>');
	expect(html).not.toMatch(/<span[^>]*data-seg-idx="1"/);
});
