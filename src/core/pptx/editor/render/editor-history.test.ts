import { describe, expect, it } from 'vitest';

import { EditorHistory } from './editor-history';

/** Walk every state undo, then redo, can reach from `current`, in the order they are reached. */
function reachable(history: EditorHistory<string>, current: string): string[] {
	const seen: string[] = [];
	let state = current;
	for (let result = history.undo(state); result; result = history.undo(state)) {
		state = result.snapshot;
		seen.push(state);
	}
	for (let result = history.redo(state); result; result = history.redo(state)) {
		state = result.snapshot;
		seen.push(state);
	}
	return seen;
}

describe('EditorHistory.forget', () => {
	it('drops the newest matching undo step and every older one', () => {
		const history = new EditorHistory<string>();
		history.record('a', 'one');
		history.record('x', 'two');
		history.record('b', 'three');
		history.record('x', 'four');
		history.record('c', 'five');

		history.forget((snapshot) => snapshot === 'x');

		expect(history.depth).toBe(1);
		expect(history.undoLabel).toBe('five');
		expect(history.undo('now')).toStrictEqual({ snapshot: 'c', label: 'five' });
		expect(history.canUndo).toBeFalsy();
	});

	it('drops the nearest matching redo step and every one beyond it', () => {
		const history = new EditorHistory<string>();
		history.record('a', 'one');
		history.record('b', 'two');
		history.record('x', 'three');
		history.record('c', 'four');
		// Back to 'b': the redo stack holds 'x' (next), 'c' (then) and 'd' (last).
		expect(history.undo('d')?.snapshot).toBe('c');
		expect(history.undo('c')?.snapshot).toBe('x');
		expect(history.undo('x')?.snapshot).toBe('b');

		history.forget((snapshot) => snapshot === 'c');

		// The undo stack did not match and is untouched.
		expect(history.depth).toBe(1);
		expect(history.redoLabel).toBe('two');
		expect(history.redo('b')).toStrictEqual({ snapshot: 'x', label: 'two' });
		expect(history.canRedo).toBeFalsy();
	});

	it('leaves no undo or redo path to a matching state', () => {
		const history = new EditorHistory<string>();
		for (const snapshot of ['a', 'x1', 'b', 'c', 'x2']) {
			history.record(snapshot, snapshot);
		}
		history.undo('d');
		history.undo('x2');

		history.forget((snapshot) => snapshot.startsWith('x'));

		expect(reachable(history, 'c')).toStrictEqual(['b', 'c']);
	});

	it('changes nothing when no step matches', () => {
		const history = new EditorHistory<string>();
		history.record('a', 'one');
		history.record('b', 'two');
		history.undo('c');

		history.forget(() => false);

		expect(history.depth).toBe(1);
		expect(history.canRedo).toBeTruthy();
		expect(reachable(history, 'b')).toStrictEqual(['a', 'b', 'c']);
	});
});
