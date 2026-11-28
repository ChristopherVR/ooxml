import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

/** Top of the paragraph box that holds `text` on the second page, in points from the content top. */
async function topOnPageTwo(name: string, text: string): Promise<number> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, createFakeMeasurer());
	const column = pages[1]?.columns[0];
	const block = column?.blocks.find(
		(item) =>
			item.kind === 'paragraph' &&
			item.lines.some((line) => line.fragments.some((fragment) => fragment.text?.includes(text))),
	);
	if (block?.kind !== 'paragraph') throw new Error(`${text} is not on page two`);
	const line = block.lines.find((item) => item.fragments.some((f) => f.text?.includes(text)));
	return ((block.yPx + (line?.yPx ?? 0)) * 3) / 4;
}

// Word 16: with the paragraph's space before of 19.4 pt and nothing after the previous paragraph,
// the text after a leading page break sits 19.4 pt below the top margin in compatibility mode 12
// (page 2 baseline 100.70 against 81.38) and at the top margin in mode 15 (81.38).
it('applies the space before again at the top of the page a leading break starts, in mode 12 (#431)', async () => {
	const bug = await topOnPageTwo('issue-431-page-break-space-before-compat12', 'Consectetur');
	const control = await topOnPageTwo(
		'issue-431-page-break-space-before-compat15-control',
		'Consectetur',
	);
	expect(control).toBeCloseTo(0, 1);
	expect(bug).toBeCloseTo(19.4, 1);
});

it('keeps the space before minus the previous paragraph space after (#431)', async () => {
	// Word: 92.66 against 81.38, so 11.4 pt of the 19.4 pt before after 8 pt of space after.
	expect(
		await topOnPageTwo('issue-431-page-break-space-before-compat12-after-previous', 'Consectetur'),
	).toBeCloseTo(11.4, 1);
});

it('keeps no space before when the break has a paragraph of its own (#431)', async () => {
	expect(
		await topOnPageTwo('issue-431-page-break-own-paragraph-compat12', 'Consectetur'),
	).toBeCloseTo(0, 1);
});
