import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { applyInlineTextDraft, reportInlineTextDraft } from './autosave-inline-draft';

const slides = [
	{
		id: 's1',
		elements: [{ id: 'title', type: 'text', text: 'Old', x: 0, y: 0, width: 10, height: 10 }],
	},
] as unknown as PptxSlide[];

describe('applyInlineTextDraft', () => {
	it('writes the draft into its element and leaves the input alone', () => {
		const next = applyInlineTextDraft(slides, { slideId: 's1', elementId: 'title', text: 'New' });
		expect(next).not.toBe(slides);
		expect((next[0]!.elements[0] as { text?: string }).text).toBe('New');
		expect((slides[0]!.elements[0] as { text?: string }).text).toBe('Old');
	});

	it('returns the same slides without a draft, an unknown element or no change', () => {
		expect(applyInlineTextDraft(slides, null)).toBe(slides);
		expect(applyInlineTextDraft(slides, { elementId: 'missing', text: 'x' })).toBe(slides);
		expect(applyInlineTextDraft(slides, { elementId: 'title', text: 'Old' })).toBe(slides);
	});
});

describe('reportInlineTextDraft', () => {
	it('hands the draft to autosave, and clears it when the editor closes', () => {
		const sink = { noteInlineDraft: vi.fn() };
		reportInlineTextDraft({ autosave: sink }, slides[0], 'title', 'Typed');
		expect(sink.noteInlineDraft).toHaveBeenLastCalledWith({
			slideId: 's1',
			elementId: 'title',
			text: 'Typed',
		});
		reportInlineTextDraft({ autosave: sink }, slides[0], null, '');
		expect(sink.noteInlineDraft).toHaveBeenLastCalledWith(null);
	});
});
