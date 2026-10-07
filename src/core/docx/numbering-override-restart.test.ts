import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { computeListLabels } from './numbering-format';
import { resolveNumberingLevel } from './numbering-parse';
import { overrideRestartFixture } from './test-support/override-restart-fixture';

// Word's documented compatibility behavior differs from the generic OOXML override rule.
// Compare generated fixtures to independently recorded native Word counters.
const reference = JSON.parse(
	await readFile(
		new URL('./__fixtures__/numbering-override-restart/native-reference.json', import.meta.url),
		'utf8',
	),
) as {
	cases: {
		name: string;
		abstractRestart: number | null;
		overrideRestart: number;
		labels: string[];
	}[];
};

for (const item of reference.cases) {
	it(`uses Word's abstract ${item.name} restart policy while preserving the ignored override`, async () => {
		const bytes = await overrideRestartFixture(
			item.abstractRestart ?? undefined,
			item.overrideRestart,
		);
		const loaded = await loadDocx(bytes);
		const catalog = loaded.model.numberingCatalog!;
		const before = structuredClone(catalog);
		expect(resolveNumberingLevel(catalog, '1', 2)?.lvlRestart).toBe(
			item.abstractRestart ?? undefined,
		);
		expect(catalog).toEqual(before);
		const labels = computeListLabels(loaded.model);
		expect(loaded.model.blocks.map((p) => labels.get(p.id)?.text)).toEqual(item.labels);
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Missing paragraph');
		paragraph.runs[0]!.text = 'Edited';
		const saved = await loaded.save();
		const source = await JSZip.loadAsync(bytes);
		const output = await JSZip.loadAsync(saved);
		expect(await output.file('word/numbering.xml')!.async('string')).toBe(
			await source.file('word/numbering.xml')!.async('string'),
		);
		const again = (await loadDocx(saved)).model;
		expect(again.numberingCatalog!.nums['1']!.levelOverrides![2]!.lvl!.lvlRestart).toBe(
			item.overrideRestart,
		);
		const reloadedLabels = computeListLabels(again);
		expect(again.blocks.map((p) => reloadedLabels.get(p.id)?.text)).toEqual(item.labels);
	});
}
