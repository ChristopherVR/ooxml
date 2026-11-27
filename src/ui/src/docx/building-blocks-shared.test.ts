import type { GlossaryEntry } from 'ooxml-core/docx';
import { describe, expect, it, vi } from 'vitest';
import { SharedBuildingBlocks, type WordBuildingBlocksStore } from './building-blocks-shared';

const entry = (name: string): GlossaryEntry => ({
	name,
	gallery: 'docParts',
	category: 'General',
	blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: name }] }],
});

function memoryStore(initial?: unknown): WordBuildingBlocksStore & { value: unknown } {
	const store = {
		value: initial,
		read: () => store.value,
		write: (value: unknown) => {
			store.value = value;
		},
	};
	return store;
}

describe('shared building blocks state', () => {
	it('reads the store when it is set and writes saved and removed entries', async () => {
		const changed = vi.fn();
		const state = new SharedBuildingBlocks(changed);
		const store = memoryStore({ version: 1, entries: [entry('Greeting')] });
		state.store = store;
		await state.load();
		expect(state.entries().map((item) => item.name)).toEqual(['Greeting']);
		await state.save(entry('Sign-off'));
		expect((store.value as { entries: GlossaryEntry[] }).entries).toHaveLength(2);
		await state.remove('docParts\u0000General\u0000Greeting');
		expect(state.entries().map((item) => item.name)).toEqual(['Sign-off']);
		expect(changed).toHaveBeenCalled();
	});

	it('ignores a read that finishes after the store was replaced', async () => {
		let release: (value: unknown) => void = () => {};
		const slow: WordBuildingBlocksStore = {
			read: () => new Promise((resolve) => (release = resolve)),
			write: () => {},
		};
		const state = new SharedBuildingBlocks(() => {});
		state.store = slow;
		state.store = memoryStore({ version: 1, entries: [entry('Current')] });
		await state.load();
		release({ version: 1, entries: [entry('Stale')] });
		await Promise.resolve();
		expect(state.entries().map((item) => item.name)).toEqual(['Current']);
	});

	it('treats a failing read as no entries', async () => {
		const state = new SharedBuildingBlocks(() => {});
		state.store = { read: () => Promise.reject(new Error('blocked')), write: () => {} };
		await state.load();
		expect(state.entries()).toEqual([]);
	});
});
