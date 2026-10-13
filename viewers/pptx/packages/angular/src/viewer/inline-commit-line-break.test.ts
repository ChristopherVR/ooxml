/**
 * Regression: committing an inline edit turned every `a:br` of the edited body
 * into a paragraph break, so the paragraph's later lines came back as new
 * paragraphs without its `a:pPr` and spacing appeared between its lines.
 * Composed the way the canvas controller commits: shared patch, then
 * `EditorStateService.updateElement`.
 */
import { buildInlineTextCommitPatch } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { EditorStateService } from './editor-state.service';

const target = {
	type: 'text',
	id: 'a',
	name: '',
	x: 0,
	y: 0,
	width: 200,
	height: 80,
	text: 'First line\nSecond line\nNext paragraph',
	textSegments: [
		{ text: 'First line', style: {}, paragraphProperties: { paragraphSpacingAfter: 5 } },
		{ text: '\n', style: {}, isLineBreak: true },
		{ text: 'Second line', style: {} },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{ text: 'Next paragraph', style: {} },
	],
} as unknown as PptxElement;

describe('inline commit with a soft line break', () => {
	it('keeps the a:br inside its paragraph when a character is typed', () => {
		const svc = new EditorStateService();
		svc.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [target] } as PptxSlide]);
		const patch = buildInlineTextCommitPatch(target, 'First lineX\nSecond line\nNext paragraph');
		svc.updateElement(0, 'a', patch!);

		const segments = (svc.slides()[0].elements[0] as { textSegments: TextSegment[] }).textSegments;
		expect(segments.filter((s) => s.isLineBreak)).toHaveLength(1);
		expect(segments.filter((s) => s.isParagraphBreak)).toHaveLength(1);
	});
});
