import { DEFAULT_EDITOR_HISTORY_DEPTH } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import type { ViewerState } from '../state';
import { createEditorOps } from './editor-operations';

function setup() {
	const slide: PptxSlide = { id: 's1', rId: 'rId1', slideNumber: 1, elements: [] };
	const store = createStore<ViewerState>({
		...createInitialViewerState(),
		slides: [slide],
		editable: true,
	});
	return createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
}

describe('vanilla editor history', () => {
	it('uses the shared history controller and its one default depth', () => {
		const ops = setup();
		expect(ops.history.getSnapshot().maxDepth).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
		for (let i = 0; i < DEFAULT_EDITOR_HISTORY_DEPTH + 5; i++) {
			ops.pushHistory(`edit ${i}`);
		}
		expect(ops.history.getSnapshot().depth).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
	});

	it('publishes canUndo and the labels through the controller snapshot', () => {
		const ops = setup();
		const listener = vi.fn();
		ops.history.subscribe(listener);
		ops.pushHistory('Move');
		expect(listener).toHaveBeenCalled();
		expect(ops.canUndo()).toBe(true);
		expect(ops.undoLabel()).toBe('Move');
		ops.undo();
		expect(ops.canRedo()).toBe(true);
		expect(ops.redoLabel()).toBe('Move');
	});
});
