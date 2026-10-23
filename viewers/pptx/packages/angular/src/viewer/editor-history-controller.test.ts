import { DEFAULT_EDITOR_HISTORY_DEPTH } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { EditorStateService } from './editor-state.service';

function deck(): PptxSlide[] {
	const shape = { type: 'shape', id: 'a', x: 0, y: 0, width: 10, height: 10 } as PptxElement;
	return [{ id: 's1', rId: 's1', slideNumber: 1, elements: [shape] } as PptxSlide];
}

describe('angular editor history (shared controller)', () => {
	it('keeps at most the one shared default depth of undo steps', () => {
		const svc = new EditorStateService();
		svc.setSlides(deck());
		for (let i = 0; i <= DEFAULT_EDITOR_HISTORY_DEPTH + 4; i++) {
			svc.updateElement(0, 'a', { x: i + 1 });
		}
		let undone = 0;
		while (svc.canUndo()) {
			svc.undo();
			undone += 1;
		}
		expect(undone).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
	});

	it('mirrors canUndo, canRedo and the labels from the controller signal', () => {
		const svc = new EditorStateService();
		svc.setSlides(deck());
		expect(svc.canUndo()).toBe(false);
		svc.updateElement(0, 'a', { x: 5 });
		expect(svc.canUndo()).toBe(true);
		svc.undo();
		expect(svc.canRedo()).toBe(true);
		svc.setHistoryDepth(3);
		expect(svc.canRedo()).toBe(true);
	});
});
