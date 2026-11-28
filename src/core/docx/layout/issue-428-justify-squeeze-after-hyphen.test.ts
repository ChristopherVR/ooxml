import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createEstimatedMeasurer } from '../pdf-export/font-measurer';
import { layoutDocumentModel } from './index';
import { justifiedSqueezeAfterDashPx } from './paragraph-layout';

// Issue #428: six justified Times New Roman 12 pt body paragraphs, the lead word one "i" longer
// each time, so `well-known` ends 2.02 pt inside the margin, then 1.32, 4.65, 7.98, 11.32 and
// 14.65 pt past it. The files were rebuilt from the issue's XML and saved by Word 16; the expected
// first lines are Word's own PDF of them. Word squeezes the 13 spaces before `well-` to keep
// `known`, up to a quarter of their width (9.75 pt), as it does for a word after a space.

/** The last word (or word part) of each paragraph's first line. */
async function firstLineEnds(name: string): Promise<string[]> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createEstimatedMeasurer()).pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? [
							(block.lines[0]?.fragments ?? [])
								.map((fragment) => fragment.text)
								.join('')
								.trim()
								.split(' ')
								.at(-1) ?? '',
						]
					: [],
			),
		),
	);
}

describe('a justified line squeezes to keep the part of a word after a hyphen (#428)', () => {
	it('keeps `well-known` up to 7.98 pt over and breaks after `well-` from 11.32 pt, as Word', async () => {
		expect(await firstLineEnds('issue-428-justify-squeeze-after-hyphen')).toEqual([
			'well-known',
			'well-known',
			'well-known',
			'well-known',
			'well-',
			'well-',
		]);
	});

	it('keeps `known` after a space in paragraphs 1 to 5 of the control, as Word', async () => {
		expect(await firstLineEnds('issue-428-justify-squeeze-after-space-control')).toEqual([
			'known',
			'known',
			'known',
			'known',
			'known',
			'well',
		]);
	});

	it('gives the part after a dash 0.3325 of its width, up to a quarter of the spaces', () => {
		// Word 16: `ab` (11.33 pt) after `well-` keeps 3.75 to 3.85 pt; 13 spaces of 3 pt cap it at 9.75.
		expect(justifiedSqueezeAfterDashPx(39, 11.328)).toBeCloseTo(3.767, 3);
		expect(justifiedSqueezeAfterDashPx(39, 32.664)).toBeCloseTo(9.75, 3);
		expect(justifiedSqueezeAfterDashPx(0, 10)).toBe(0);
	});
});
