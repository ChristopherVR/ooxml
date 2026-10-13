/**
 * Regression: an inline edit turned every `a:br` of the edited body into a
 * paragraph break, so a saved deck split one paragraph into several `a:p`
 * without the paragraph's `a:pPr` (its spacing then showed between lines).
 * Remap through the React surface and save, then read the slide XML back.
 */
import JSZip from 'jszip';
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { createTextElement, PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { remapTextToSegments } from './remap-text';

const SEGMENTS: TextSegment[] = [
	{ text: 'First line', style: {}, paragraphProperties: { paragraphSpacingAfter: 5 } },
	{ text: '\n', style: {}, isLineBreak: true },
	{ text: 'Second line', style: {} },
	{ text: '\n', style: {}, isParagraphBreak: true },
	{ text: 'Next paragraph', style: {} },
];

describe('remapTextToSegments with a soft line break (react)', () => {
	it('saves the edited a:br inside its paragraph', async () => {
		const text = 'First lineX\nSecond line\nNext paragraph';
		const segments = remapTextToSegments(text, SEGMENTS, {});
		expect(segments.filter((s) => s.isLineBreak)).toHaveLength(1);
		expect(segments.filter((s) => s.isParagraphBreak)).toHaveLength(1);

		const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
		const box = {
			...createTextElement('placeholder', { x: 10, y: 10, width: 300, height: 80 }),
			text,
			textSegments: segments,
		} as PptxElement;
		data.slides[0].elements = [box];
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const paragraphs = xml.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? [];
		const first = paragraphs.find((p) => p.includes('First lineX'));
		expect(first).toContain('<a:br');
		expect(first).toContain('Second line');
	});
});
