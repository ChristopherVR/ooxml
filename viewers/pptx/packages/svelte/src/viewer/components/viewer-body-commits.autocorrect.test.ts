import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import type { EditorState } from '../editor/editor-state.svelte';
import { createEditCommits } from './viewer-body-commits';

/** A table cell edit hands AutoCorrect the cell's text from before the edit. */
describe('createEditCommits commitTableCell autocorrect scope', () => {
	it('passes the previous cell text so only the typed part is corrected', () => {
		const table = {
			type: 'table',
			id: 'tbl',
			x: 0,
			y: 0,
			width: 100,
			height: 40,
			tableData: { rows: [{ cells: [{ text: 'Before' }] }], columnWidths: [1] },
		} as unknown as PptxElement;
		const applyElementPatch = vi.fn();
		const transformCommittedText = vi.fn((text: string) => text);
		const editor = {
			activeElements: [table],
			applyElementPatch,
			transformCommittedText,
		} as unknown as EditorState;

		createEditCommits(editor).commitTableCell('tbl', 0, 0, 'Before and after');

		expect(transformCommittedText).toHaveBeenCalledWith('Before and after', 'Before');
		expect(applyElementPatch).toHaveBeenCalledOnce();
	});
});
