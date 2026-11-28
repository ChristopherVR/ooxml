import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

/** The text of the list marker fragments of each paragraph. */
async function markers(name: string): Promise<string[]> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? [
							block.lines
								.flatMap((line) => line.fragments)
								.filter((fragment) => fragment.sourceStart === fragment.sourceEnd)
								.map((fragment) => fragment.text)
								.join(''),
						]
					: [],
			),
		),
	);
}

// Word 16: a soft hyphen that is the whole level text of a bullet is drawn as a hyphen, on every
// item including an empty one; it is not an optional break point there.
it('draws a soft hyphen bullet on every item (#433)', async () => {
	const bug = await markers('issue-433-bullet-soft-hyphen');
	const control = await markers('issue-433-bullet-hyphen-minus-control');
	expect(control).toEqual(['-', '-', '-', '-']);
	expect(bug).toHaveLength(4);
	expect(bug).toEqual(control);
});
