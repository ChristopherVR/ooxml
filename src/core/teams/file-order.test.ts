import { describe, expect, it } from 'vitest';
import { sortFiles } from './file-order';
import type { FileEntry } from './view';

const entries: FileEntry[] = [
	{ name: 'Budget10.xlsx', author: 'Ada', ts: 20, kind: 'xlsx', messageId: 'one' },
	{ name: 'budget2.xlsx', author: 'Bob', ts: 30, kind: 'xlsx', messageId: 'two' },
	{ name: 'budget2.xlsx', author: 'Ada', ts: 10, kind: 'xlsx', messageId: 'three' },
];

describe('shared file ordering', () => {
	it('uses natural names, stable ties and does not mutate source entries', () => {
		const sorted = sortFiles(entries, 'name', 'ascending');
		expect(sorted).toEqual([entries[1], entries[2], entries[0]]);
		expect(sorted[0]).toBe(entries[1]);
		expect(entries.map((entry) => entry.messageId)).toEqual(['one', 'two', 'three']);
	});
	it('sorts shared dates and authors in both directions', () => {
		expect(sortFiles(entries, 'ts', 'descending')).toEqual([entries[1], entries[0], entries[2]]);
		expect(sortFiles(entries, 'ts', 'ascending')).toEqual([entries[2], entries[0], entries[1]]);
		expect(sortFiles(entries, 'author', 'ascending')).toEqual([entries[0], entries[2], entries[1]]);
		expect(sortFiles(entries, 'author', 'descending')).toEqual([
			entries[1],
			entries[0],
			entries[2],
		]);
	});
});
