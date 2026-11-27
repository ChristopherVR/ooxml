import { describe, expect, it } from 'vitest';
import type { GlossaryEntry } from './glossary-document';
import {
	MAX_SHARED_BUILDING_BLOCKS,
	parseSharedBuildingBlocks,
	sharedBuildingBlocksValue,
} from './glossary-shared';

const entry = (name: string, text = name): GlossaryEntry => ({
	name,
	gallery: 'docParts',
	category: 'General',
	blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text }] }],
});

describe('shared building blocks', () => {
	it('round trips entries through the stored value, without file markers', () => {
		const value = sharedBuildingBlocksValue([{ ...entry('Sign-off'), loaded: true }]);
		expect(value.entries[0]).not.toHaveProperty('loaded');
		expect(parseSharedBuildingBlocks(JSON.parse(JSON.stringify(value)))).toEqual([
			entry('Sign-off'),
		]);
		expect(parseSharedBuildingBlocks(JSON.stringify(value))).toEqual([entry('Sign-off')]);
	});

	it('keeps the newest copy of an entry with the same name, gallery and category', () => {
		const value = sharedBuildingBlocksValue([entry('A', 'old'), entry('A', 'new')]);
		expect(value.entries).toHaveLength(1);
		expect(value.entries[0]?.blocks[0]).toMatchObject({ runs: [{ text: 'new' }] });
	});

	it('reads a damaged, foreign or oversized value as no entries', () => {
		expect(parseSharedBuildingBlocks(undefined)).toEqual([]);
		expect(parseSharedBuildingBlocks('{not json')).toEqual([]);
		expect(parseSharedBuildingBlocks({ version: 2, entries: [entry('A')] })).toEqual([]);
		expect(
			parseSharedBuildingBlocks({
				version: 1,
				entries: [
					{ name: 'No blocks', gallery: 'docParts', category: 'General', blocks: [] },
					{ ...entry('Bad run'), blocks: [{ type: 'paragraph', id: 'p', runs: [{}] }] },
					{
						...entry('Picture'),
						blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: '', image: {} }] }],
					},
					entry('Kept'),
				],
			}).map((item) => item.name),
		).toEqual(['Kept']);
	});

	it('keeps at most the limit', () => {
		const many = Array.from({ length: MAX_SHARED_BUILDING_BLOCKS + 5 }, (_, i) => entry(`E${i}`));
		expect(parseSharedBuildingBlocks({ version: 1, entries: many })).toHaveLength(
			MAX_SHARED_BUILDING_BLOCKS,
		);
	});
});
