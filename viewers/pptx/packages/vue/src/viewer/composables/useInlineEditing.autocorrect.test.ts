import { DEFAULT_VIEWER_OPTIONS } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorOperations } from './useEditorOperations';
import { useInlineEditing } from './useInlineEditing';

/**
 * AutoCorrect acts on what was typed in this edit, never on text that was already in the box (a
 * URL after a sentence, or a collaborator's words).
 */
function makeElement(text: string): PptxElement {
	return {
		id: 'text-1',
		type: 'text',
		x: 0,
		y: 0,
		width: 300,
		height: 40,
		text,
		textSegments: [{ text, style: {} }],
		textStyle: {},
	} as unknown as PptxElement;
}

function commit(existing: string, typed: string): string | undefined {
	const element = makeElement(existing);
	const updateElement = vi.fn();
	const editing = useInlineEditing({
		canEdit: () => true,
		findActiveElement: (id) => (id === element.id ? element : undefined),
		ops: { updateElement } as unknown as EditorOperations,
		proofing: () => DEFAULT_VIEWER_OPTIONS.proofing,
	});
	editing.enterInlineEdit(element.id);
	editing.updateInlineText(typed);
	editing.commitInlineEdit();
	return (updateElement.mock.calls[0]?.[1] as { text?: string } | undefined)?.text;
}

describe('commitInlineEdit autocorrect scope', () => {
	const existing = 'Fine. https://example.com/a has IDs on monday.';

	it('leaves existing text alone and corrects only the typed words', () => {
		expect(commit(existing, `${existing} typed monday`)).toBe(`${existing} Typed Monday`);
	});

	it('does not rewrite the box when nothing was typed', () => {
		expect(commit(existing, `${existing}!`)).toBe(`${existing}!`);
	});
});
