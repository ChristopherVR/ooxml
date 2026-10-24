import { describe, expect, it, vi } from 'vitest';

import { DEFAULT_EDITOR_HISTORY_DEPTH } from '../editor-history';
import { DEFAULT_VIEWER_OPTIONS } from '../options/viewer-options';
import { resolveHistoryDepth } from '../options/viewer-options-apply';
import { createEditorHistoryController } from './history-controller';

describe('createEditorHistoryController', () => {
	it('defaults to the one shared depth, which is the File > Options default', () => {
		const history = createEditorHistoryController<number>();
		expect(history.getSnapshot().maxDepth).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
		expect(DEFAULT_EDITOR_HISTORY_DEPTH).toBe(100);
		expect(resolveHistoryDepth(DEFAULT_VIEWER_OPTIONS)).toBe(DEFAULT_EDITOR_HISTORY_DEPTH);
	});

	it('records, undoes and redoes with observable flags, labels and revision', () => {
		const history = createEditorHistoryController<string>();
		const listener = vi.fn();
		history.subscribe(listener);
		expect(history.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false, revision: 0 });

		history.record('a', 'Move');
		expect(history.getSnapshot()).toMatchObject({
			canUndo: true,
			canRedo: false,
			undoLabel: 'Move',
			depth: 1,
			revision: 1,
		});
		expect(history.undo('b')).toEqual({ snapshot: 'a', label: 'Move' });
		expect(history.getSnapshot()).toMatchObject({
			canUndo: false,
			canRedo: true,
			redoLabel: 'Move',
			revision: 2,
		});
		expect(history.redo('a')).toEqual({ snapshot: 'b', label: 'Move' });
		expect(history.getSnapshot().revision).toBe(3);
		expect(listener).toHaveBeenCalledTimes(3);
	});

	it('reports an unlabelled entry as undefined and ignores an empty undo', () => {
		const history = createEditorHistoryController<number>();
		expect(history.undo(1)).toBeUndefined();
		expect(history.getSnapshot().revision).toBe(0);
		history.record(1);
		expect(history.getSnapshot().undoLabel).toBeUndefined();
	});

	it('trims to a smaller maximum without dropping the newest entries', () => {
		const history = createEditorHistoryController<number>({ maxDepth: 5 });
		for (let i = 0; i < 5; i++) {
			history.record(i, `step ${i}`);
		}
		history.setMaxDepth(2);
		expect(history.getSnapshot()).toMatchObject({ depth: 2, maxDepth: 2, undoLabel: 'step 4' });
	});

	it('forgets the steps that reach a matching state and reports what is left', () => {
		const history = createEditorHistoryController<string>();
		history.record('a', 'one');
		history.record('x', 'two');
		history.record('b', 'three');
		history.record('d', 'four');
		history.undo('c');
		const listener = vi.fn();
		history.subscribe(listener);

		history.forget((snapshot) => snapshot === 'x' || snapshot === 'c');

		expect(history.getSnapshot()).toMatchObject({
			canUndo: true,
			canRedo: false,
			undoLabel: 'three',
			redoLabel: undefined,
			depth: 1,
		});
		expect(listener).toHaveBeenCalledOnce();

		// Nothing left to forget: no change, no notification.
		history.forget((snapshot) => snapshot === 'x');
		expect(listener).toHaveBeenCalledOnce();
		expect(history.undo('d')).toEqual({ snapshot: 'b', label: 'three' });
		expect(history.getSnapshot().canUndo).toBeFalsy();
	});

	it('clears both stacks', () => {
		const history = createEditorHistoryController<number>();
		history.record(1, 'x');
		history.undo(2);
		history.clear();
		expect(history.getSnapshot()).toMatchObject({ canUndo: false, canRedo: false });
	});
});
