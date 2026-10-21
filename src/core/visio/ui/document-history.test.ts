import { expect, it } from 'vitest';
import { DocumentHistory, EDIT_HISTORY_LIMITS } from './document-history';
it('owns snapshots, truncates by entry count and pins the original independently', () => {
	const bytes = new Uint8Array([1]);
	const history = new DocumentHistory(bytes);
	bytes[0] = 8;
	expect(history.current.bytes[0]).toBe(1);
	for (let i = 0; i < EDIT_HISTORY_LIMITS.maxEntries + 2; i++)
		history.append(new Uint8Array([i + 2]), []);
	expect(history.state.historyTruncated).toBe(true);
	expect(history.original.bytes[0]).toBe(1);
	let count = 0;
	while (history.undoTarget) {
		history.move(history.undoTarget);
		count++;
	}
	expect(count).toBe(EDIT_HISTORY_LIMITS.maxEntries - 1);
	expect(history.state.dirty).toBe(true);
});
it('caps aggregate retained bytes including pinned original and current', () => {
	const history = new DocumentHistory(new Uint8Array(32 * 1024 * 1024));
	for (let i = 0; i < 4; i++) history.append(new Uint8Array(32 * 1024 * 1024), []);
	expect(history.state.historyTruncated).toBe(true);
	let count = 0;
	while (history.undoTarget) {
		history.move(history.undoTarget);
		count++;
	}
	expect(count).toBe(1);
});
it('measures dirty from the last saved snapshot after markClean', () => {
	const history = new DocumentHistory(new Uint8Array([1]));
	history.append(new Uint8Array([2]), []);
	expect(history.state.dirty).toBe(true);
	history.markClean();
	expect(history.state.dirty).toBe(false);
	expect(history.export().dirty).toBe(false);
	history.append(new Uint8Array([3]), []);
	expect(history.state.dirty).toBe(true);
	history.move(history.undoTarget!);
	expect(history.state.dirty).toBe(false);
	history.move(history.undoTarget!);
	expect(history.state.dirty).toBe(true);
	expect(history.original.bytes[0]).toBe(2);
});
