import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import type { DocumentModel } from '../model';
import { loadDocx } from '../parse';
import { adaptDocumentModel } from './adapter';
import { withComputedFieldRuns } from './computed-field-runs';
import type { LayoutParagraph } from './input';

const fixture = (name: string) =>
	readFile(new URL(`../__fixtures__/computed-fields/${name}.docx`, import.meta.url));

/** Each body paragraph as the layout draws it, and the source length of its runs. */
export function drawn(model: DocumentModel) {
	const input = withComputedFieldRuns(adaptDocumentModel(model), model);
	return input.sections
		.flatMap((section) => section.blocks)
		.filter((block): block is LayoutParagraph => block.kind === 'paragraph')
		.map((paragraph) => {
			const runs = paragraph.runs.filter((run) => !run.synthetic && !run.hidden);
			return {
				text: runs.map((run) => run.text).join(''),
				length: runs.reduce((sum, run) => sum + (run.sourceLength ?? run.text.length), 0),
			};
		});
}

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
