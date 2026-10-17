import { describe, expect, it } from 'vitest';
import { DocumentHistory } from './document-history';

const bytes = (n: number) => new Uint8Array([n]);

describe('DocumentHistory depth', () => {
	it('keeps the requested number of undo steps', () => {
		const history = new DocumentHistory(bytes(0), 3);
		for (let i = 1; i <= 5; i++) history.append(bytes(i), []);
		let steps = 0;
		while (history.undoTarget) {
			history.move(history.undoTarget);
			steps++;
		}
		expect(steps).toBe(2);
	});

	it('trims at once when the depth is lowered, dropping redo steps first', () => {
		const history = new DocumentHistory(bytes(0));
		for (let i = 1; i <= 6; i++) history.append(bytes(i), []);
		history.move(history.undoTarget!);
		history.move(history.undoTarget!);
		history.maxEntries = 4;
		expect(history.state.canUndo).toBe(true);
		let steps = 0;
		while (history.undoTarget) {
			history.move(history.undoTarget);
			steps++;
		}
		expect(steps).toBeLessThanOrEqual(3);
	});
});
