import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import { computeListLabels, formatListNumber } from './numbering-format';
import { hexNumberingFixture, hexCases } from './test-support/hex-numbering-fixture';

const reference = JSON.parse(
	await readFile(
		new URL('./__fixtures__/hex-numbering/native-reference.json', import.meta.url),
		'utf8',
	),
) as { cases: { name: string; template: string; value: number; label: string }[] };

it.each(reference.cases.filter((item) => item.template === '%1'))(
	'matches native Word hex label $value -> $label',
	({ value, label }) => {
		expect(formatListNumber('hex', value)).toBe(label);
	},
);

it('matches native hex lists and retains numbering definitions after text editing', async () => {
	const bytes = await hexNumberingFixture();
	const loaded = await loadDocx(bytes);
	expect(reference.cases.map(({ value, template }) => ({ value, template }))).toEqual(hexCases);
	const assertLabels = (model: typeof loaded.model) => {
		const labels = computeListLabels(model);
		expect(model.blocks.map((block) => labels.get(block.id)?.text)).toEqual(
			reference.cases.map((item) => item.label),
		);
	};
	assertLabels(loaded.model);
	const paragraph = loaded.model.blocks[0]!;
	if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
	paragraph.runs[0]!.text = 'Edited';
	const saved = await loaded.save();
	const inputZip = await JSZip.loadAsync(bytes);
	const outputZip = await JSZip.loadAsync(saved);
	expect(await outputZip.file('word/numbering.xml')!.async('string')).toBe(
		await inputZip.file('word/numbering.xml')!.async('string'),
	);
	assertLabels((await loadDocx(saved)).model);
});
