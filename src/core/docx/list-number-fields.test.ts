import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { computeListLabels } from './numbering-format';
import { loadDocx } from './parse';
import { drawn } from './test-support/computed-fields';

// Word 16 inserted and updated these fields (`__fixtures__/computed-fields/`); it saves them
// without results, so what it shows was read from its PDF export into native-reference.json.
const folder = new URL('./__fixtures__/computed-fields/', import.meta.url);
const reference = JSON.parse(await readFile(new URL('native-reference.json', folder), 'utf8')) as {
	documents: Record<string, string[]>;
};

const normal = (text: string) => text.replace(/\s+/g, ' ').trim();

describe('LISTNUM', () => {
	it.each(Object.keys(reference.documents).filter((name) => name.startsWith('list-number')))(
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

	it('moves the list of its paragraph on, so later list labels count it', async () => {
		const { model } = await loadDocx(await readFile(new URL('list-number-1.docx', folder)));
		const labels = computeListLabels(model);
		const shown = model.blocks.flatMap((block) => labels.get(block.id)?.text ?? []);
		expect(shown).toEqual(['1.', '4.', '6.']);
	});
});
