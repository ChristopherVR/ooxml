import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { computeListLabels } from './numbering-format';
import { resolveNumberingLevel } from './numbering-parse';
import {
	numberingStartFixture,
	type NumberingStartCase,
} from './test-support/numbering-start-fixture';

const reference = JSON.parse(
	await readFile(
		new URL('./__fixtures__/numbering-override-start/native-reference.json', import.meta.url),
		'utf8',
	),
) as { cases: (NumberingStartCase & { labels: string[] })[] };

for (const item of reference.cases) {
	it(`matches native Word start values for ${item.name} without rewriting source numbering`, async () => {
		const bytes = await numberingStartFixture(item);
		const loaded = await loadDocx(bytes);
		const catalog = loaded.model.numberingCatalog!;
		const original = structuredClone(catalog);
		const labels = computeListLabels(loaded.model);
		expect(loaded.model.blocks.map((block) => labels.get(block.id)?.text)).toEqual(item.labels);
		expect(catalog).toEqual(original);
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Missing paragraph');
		paragraph.runs[0]!.text = 'Edited';
		const saved = await loaded.save();
		const before = await JSZip.loadAsync(bytes);
		const after = await JSZip.loadAsync(saved);
		expect(await after.file('word/numbering.xml')!.async('string')).toBe(
			await before.file('word/numbering.xml')!.async('string'),
		);
		const reloaded = (await loadDocx(saved)).model;
		const again = computeListLabels(reloaded);
		expect(reloaded.blocks.map((block) => again.get(block.id)?.text)).toEqual(item.labels);
	});
}

it('keeps an omitted start omitted when adding another instance with its imported full level', async () => {
	const item = reference.cases.find((item) => item.name === 'default-full-missing-start')!;
	const loaded = await loadDocx(await numberingStartFixture(item));
	const catalog = loaded.model.numberingCatalog!;
	const definition = catalog.nums['1']!;
	expect(definition.levelOverrides![2]!.lvl!.startWasOmitted).toBe(true);
	catalog.nums['2'] = { ...structuredClone(definition), id: '2' };
	const saved = await loaded.save();
	const reloaded = await loadDocx(saved);
	const created = reloaded.model.numberingCatalog!.nums['2']!.levelOverrides![2]!.lvl!;
	expect(created.startWasOmitted).toBe(true);
	expect(resolveNumberingLevel(reloaded.model.numberingCatalog!, '2', 2)?.start).toBe(7);
});

it('treats a programmatic full-level start as explicit without requiring parser provenance', async () => {
	const item = reference.cases.find((item) => item.name === 'default-full-conflict')!;
	const loaded = await loadDocx(await numberingStartFixture(item));
	const catalog = loaded.model.numberingCatalog!;
	const level = catalog.nums['1']!.levelOverrides![2]!.lvl!;
	delete level.startWasOmitted;
	level.start = 1;
	expect(resolveNumberingLevel(catalog, '1', 2)?.start).toBe(1);
});

it('writes and resolves an authored start changed from the omitted-source fallback', async () => {
	const item = reference.cases.find((item) => item.name === 'default-full-missing-start')!;
	const loaded = await loadDocx(await numberingStartFixture(item));
	const catalog = loaded.model.numberingCatalog!;
	const created = { ...structuredClone(catalog.nums['1']!), id: '2' };
	created.levelOverrides![2]!.lvl!.start = 5;
	expect(created.levelOverrides![2]!.lvl!.startWasOmitted).toBe(true);
	catalog.nums['2'] = created;
	expect(resolveNumberingLevel(catalog, '2', 2)?.start).toBe(5);
	const reloaded = await loadDocx(await loaded.save());
	const saved = reloaded.model.numberingCatalog!.nums['2']!.levelOverrides![2]!.lvl!;
	expect(saved.start).toBe(5);
	expect(saved.startWasOmitted).toBeUndefined();
	expect(resolveNumberingLevel(reloaded.model.numberingCatalog!, '2', 2)?.start).toBe(5);
});

it('keeps omitted abstract starts and labels stable when copying a definition into another list', async () => {
	const item = reference.cases.find((item) => item.name === 'abstract-default-missing')!;
	const loaded = await loadDocx(await numberingStartFixture(item));
	const catalog = loaded.model.numberingCatalog!;
	const abstract = catalog.abstractNums['0']!;
	expect(abstract.levels[2]!.startWasOmitted).toBe(true);
	catalog.abstractNums['1'] = { ...structuredClone(abstract), id: '1' };
	catalog.nums['2'] = { id: '2', abstractNumId: '1' };
	for (const block of loaded.model.blocks) {
		if (block.type === 'paragraph' && block.numbering) block.numbering.numId = 2;
	}
	const before = computeListLabels(loaded.model);
	expect(loaded.model.blocks.map((block) => before.get(block.id)?.text)).toEqual(item.labels);
	const reloaded = (await loadDocx(await loaded.save())).model;
	expect(reloaded.numberingCatalog!.abstractNums['1']!.levels[2]!.startWasOmitted).toBe(true);
	const after = computeListLabels(reloaded);
	expect(reloaded.blocks.map((block) => after.get(block.id)?.text)).toEqual(item.labels);
});
