import { DEFAULT_EDITOR_HISTORY_DEPTH } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';
import { flushSync } from 'svelte';
import { describe, expect, it } from 'vitest';

import { EditorState } from './editor-state.svelte';

function makeEditor(): EditorState {
	const editor = new EditorState({
		getCurrent: () => 0,
		getHandler: () => ({}) as unknown as PptxHandler,
	});
	editor.editable = true;
	return editor;
}

describe('svelte editor history', () => {
	it('uses the shared history controller and its one default depth', () => {
		const editor = makeEditor();
		expect(editor.history.controller.getSnapshot().maxDepth).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
		for (let i = 0; i < DEFAULT_EDITOR_HISTORY_DEPTH + 5; i++) {
			editor.pushHistory(`edit ${i}`);
		}
		expect(editor.history.controller.getSnapshot().depth).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
	});

	it('keeps canUndo reactive through the controller adapter, with no manual sync', () => {
		const editor = makeEditor();
		const seen: boolean[] = [];
		const cleanup = $effect.root(() => {
			$effect(() => {
				seen.push(editor.history.canUndo);
			});
		});
		flushSync();
		editor.history.record(editor.snapshot(), 'Move');
		flushSync();
		expect(seen).toEqual([false, true]);
		expect(editor.history.undoLabel).toBe('Move');
		cleanup();
	});

	it('keeps the newest entries when the maximum shrinks', () => {
		const editor = makeEditor();
		for (let i = 0; i < 5; i++) {
			editor.pushHistory(`edit ${i}`);
		}
		editor.setHistoryDepth(3);
		expect(editor.history.controller.getSnapshot()).toMatchObject({
			depth: 3,
			undoLabel: 'edit 4',
		});
	});
});
