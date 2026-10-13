// @vitest-environment jsdom
/**
 * Regression: the inline editor of an `a:bodyPr/@wrap="none"` body showed all
 * of its lines joined on one line. The surface takes the renderer's block
 * style, which set `white-space: nowrap` for a no-wrap body; `nowrap` also
 * collapses the "\n" breaks the surface holds the body as, so the paragraphs
 * and `a:br` lines ran together while editing (PowerPoint keeps them).
 */
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { openInlineEditor, readEditableText } from './inline-text-editor';

function noWrapElement(): PptxElement {
	return {
		id: 'el-1',
		type: 'text',
		x: 0,
		y: 0,
		width: 200,
		height: 80,
		text: 'First line\nSecond line\nNext paragraph',
		textStyle: { textWrap: 'none' },
		textSegments: [
			{ text: 'First line', style: {} },
			{ text: '\n', style: {}, isLineBreak: true },
			{ text: 'Second line', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'Next paragraph', style: {} },
		],
	} as unknown as PptxElement;
}

describe('openInlineEditor on a wrap="none" body', () => {
	it('keeps the line breaks visible: white-space is pre, never nowrap', () => {
		const overlayRoot = document.createElement('div');
		document.body.append(overlayRoot);
		const session = openInlineEditor({
			doc: document,
			overlayRoot,
			box: { x: 0, y: 0, width: 200, height: 80, rotation: 0 },
			scale: 1,
			element: noWrapElement(),
			onCommit: vi.fn(),
			onClose: vi.fn(),
		});

		expect(session.el.style.whiteSpace).toBe('pre');
		expect(readEditableText(session.el)).toBe('First line\nSecond line\nNext paragraph');

		session.cancel();
		overlayRoot.remove();
	});
});
