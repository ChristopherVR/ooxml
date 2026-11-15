import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { autoNumberResult } from './auto-number-fields';
import { drawn } from './test-support/computed-fields';
import { refreshComputedFields } from './computed-fields';
import { loadDocx } from './parse';

// Word 16 inserted and updated these fields (`__fixtures__/computed-fields/`); it saves them
// without results, so what it shows was read from its PDF export into native-reference.json.
const folder = new URL('./__fixtures__/computed-fields/', import.meta.url);
const reference = JSON.parse(await readFile(new URL('native-reference.json', folder), 'utf8')) as {
	documents: Record<string, string[]>;
};

const normal = (text: string) => text.replace(/\s+/g, ' ').trim();

describe('AUTONUM, AUTONUMLGL and AUTONUMOUT', () => {
	it.each(Object.keys(reference.documents).filter((name) => name.startsWith('auto-number')))(
		'draws %s as Word does',
		async (name) => {
			const { model } = await loadDocx(await readFile(new URL(`${name}.docx`, folder)));
			expect(
				drawn(model)
					.map((item) => normal(item.text))
					.filter(Boolean),
			).toEqual(reference.documents[name]);
		},
	);

	it('saves fields without a result as Word does, and brings a saved result up to date', async () => {
		const loaded = await loadDocx(await readFile(new URL('auto-number-1.docx', folder)));
		const stale = {
			type: 'paragraph' as const,
			id: 'stale',
			runs: [{ text: '9.', field: { instr: ' AUTONUM ', simple: true } }, { text: ' z' }],
		};
		const saved = await loaded.save({
			...loaded.model,
			blocks: [...loaded.model.blocks.slice(0, 2), stale, ...loaded.model.blocks.slice(2)],
		});
		const { model } = await loadDocx(saved);
		expect(
			drawn(model)
				.map((item) => normal(item.text))
				.slice(0, 4),
		).toEqual(['1. a', '2. b', '3. z', '4: c']);
		const third = model.blocks[2];
		expect(third?.type === 'paragraph' && third.runs.map((run) => run.text).join('')).toBe('3. z');
		expect(
			model.blocks[0]?.type === 'paragraph' &&
				model.blocks[0].runs.some((run) => run.fieldChar === 'separate'),
		).toBe(false);
	});

	it('reads separators, formats and levels', () => {
		const body = { level: 3, heading: false, counts: [2, 1, 4] };
		const heading = { level: 2, heading: true, counts: [2, 1] };
		expect(autoNumberResult('AUTONUM', body)).toBe('4.');
		expect(autoNumberResult('AUTONUM \\s "--" \\* ROMAN', body)).toBe('IV-');
		expect(autoNumberResult('AUTONUMLGL \\e', body)).toBe('2.1.4');
		expect(autoNumberResult('AUTONUMOUT', heading)).toBe('A.');
		expect(autoNumberResult('AUTONUMOUT \\s :', body)).toBe('4.');
		expect(autoNumberResult('AUTONUM \\# 0', body)).toBeUndefined();
		expect(autoNumberResult('SEQ x', body)).toBeUndefined();
	});

	it('leaves a field without a result slot alone when results are refreshed', async () => {
		const { model } = await loadDocx(await readFile(new URL('auto-number-2.docx', folder)));
		expect(refreshComputedFields(model, { body: true })).toBe(model);
	});
});
