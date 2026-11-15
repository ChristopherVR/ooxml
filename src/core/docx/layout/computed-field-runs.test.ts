import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../parse';
import { drawn } from '../test-support/computed-fields';

const fixture = (name: string) =>
	readFile(new URL(`../__fixtures__/computed-fields/${name}.docx`, import.meta.url));

describe('computed field runs', () => {
	it('draws DOCPROPERTY results from the properties, keeping source lengths', async () => {
		const { model } = await loadDocx(await fixture('doc-property'));
		const before = drawn(model);
		const after = drawn({ ...model, properties: { ...model.properties, title: 'Edited' } });
		expect(before[0]?.text).toBe('my report title');
		expect(after[0]).toEqual({ text: 'Edited', length: before[0]!.length });
		expect(after[1]?.text).toBe('EDITED');
		const others = (items: typeof before) =>
			items
				.filter((_, index) => !/report title/i.test(before[index]!.text))
				.map((item) => item.text);
		expect(others(after)).toEqual(others(before));
	});
});
