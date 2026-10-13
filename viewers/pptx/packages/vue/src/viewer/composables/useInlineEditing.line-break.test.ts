/**
 * Regression: committing an inline edit turned every `a:br` of the edited body
 * into a paragraph break, so the paragraph's later lines came back as new
 * paragraphs without its `a:pPr` and spacing appeared between its lines.
 */
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorOperations } from './useEditorOperations';
import { useInlineEditing } from './useInlineEditing';

function element(): PptxElement {
	return {
		id: 'text-1',
		type: 'text',
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
}

describe('commitInlineEdit with a soft line break', () => {
	it('keeps the a:br inside its paragraph when a character is typed', () => {
		const target = element();
		const updateElement = vi.fn();
		const editing = useInlineEditing({
			canEdit: () => true,
			findActiveElement: (id) => (id === target.id ? target : undefined),
			ops: { updateElement } as unknown as EditorOperations,
		});

		editing.enterInlineEdit(target.id);
		editing.updateInlineText('First lineX\nSecond line\nNext paragraph');
		editing.commitInlineEdit();

		const segments = updateElement.mock.calls[0][1].textSegments as TextSegment[];
		expect(segments.filter((s) => s.isLineBreak)).toHaveLength(1);
		expect(segments.filter((s) => s.isParagraphBreak)).toHaveLength(1);
	});
});
