import { applyAutoCorrectToEdit, DEFAULT_VIEWER_OPTIONS } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createEditorOps } from './editor-operations';

/**
 * AutoCorrect acts on what was typed in this edit. The editor hands the committed text and the
 * text from before the edit to `transformCommittedText`, so a URL or a collaborator's words that
 * were already in the box are left exactly as they were.
 */
const EXISTING = 'Fine. https://example.com/a has IDs on monday.';

function textBox(text: string): PptxElement {
	return { type: 'text', id: 't1', x: 0, y: 0, width: 300, height: 40, text } as PptxElement;
}

function makeOps(elements: PptxElement[]) {
	const slide: PptxSlide = { id: 'a', rId: 'rId-a', slideNumber: 1, elements, notes: '' };
	const store = createStore({
		...createInitialViewerState(),
		slides: [slide],
		currentSlide: 0,
		editable: true,
	});
	const transformCommittedText = vi.fn((text: string, previous?: string) =>
		applyAutoCorrectToEdit(text, previous, DEFAULT_VIEWER_OPTIONS.proofing),
	);
	const ops = createEditorOps({
		store,
		getHandler: () => null,
		onHistoryChange: vi.fn(),
		transformCommittedText,
	});
	return { store, ops, transformCommittedText };
}

describe('createEditorOps commitInlineText autocorrect scope', () => {
	it('corrects only the typed words and leaves the existing text as it was', () => {
		const { store, ops, transformCommittedText } = makeOps([textBox(EXISTING)]);
		ops.commitInlineText('t1', `${EXISTING} typed monday`);
		expect(transformCommittedText).toHaveBeenCalledWith(`${EXISTING} typed monday`, EXISTING);
		const element = store.get().slides[0]?.elements[0] as { text?: string } | undefined;
		expect(element?.text).toBe(`${EXISTING} Typed Monday`);
	});
});

describe('createEditorOps commitTableCell autocorrect scope', () => {
	it('hands autocorrect the cell text from before the edit', () => {
		const table = {
			type: 'table',
			id: 'tbl',
			x: 0,
			y: 0,
			width: 100,
			height: 40,
			tableData: { rows: [{ cells: [{ text: EXISTING }] }], columnWidths: [1] },
		} as unknown as PptxElement;
		const { store, ops, transformCommittedText } = makeOps([table]);
		ops.commitTableCell('tbl', 0, 0, `${EXISTING} typed monday`);
		expect(transformCommittedText).toHaveBeenCalledWith(`${EXISTING} typed monday`, EXISTING);
		const cell = (
			store.get().slides[0]?.elements[0] as unknown as {
				tableData: { rows: { cells: { text: string }[] }[] };
			}
		).tableData.rows[0]?.cells[0];
		expect(cell?.text).toBe(`${EXISTING} Typed Monday`);
	});
});
