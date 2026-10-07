import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { authoredNumberingStartFixture } from './test-support/authored-numbering-start-fixture';
import type { NumberingStartCase } from './test-support/numbering-start-fixture';
import { loadDocx } from './parse';
import { computeListLabels } from './numbering-format';

const reference = JSON.parse(
	await readFile(
		new URL('./__fixtures__/numbering-override-edit/native-reference.json', import.meta.url),
		'utf8',
	),
) as {
	cases: (NumberingStartCase & { labels: string[] })[];
};
for (const item of reference.cases) {
	it(`exports ${item.name} with counters matching independently reopened native Word`, async () => {
		const model = (await loadDocx(await authoredNumberingStartFixture(item))).model;
		const labels = computeListLabels(model);
		expect(model.blocks.map((block) => labels.get(block.id)?.text)).toEqual(item.labels);
	});
}
