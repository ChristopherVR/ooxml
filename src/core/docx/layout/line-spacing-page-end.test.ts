import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { centuryGothicMeasurer } from './test-support/century-gothic-measurer';

// Word 16.0.20625's PDF of each #383 sample: Century Gothic 10 pt, 8 pt after, A4 with 1417 twip
// top and bottom margins. At 1.08 lines (w:line 259) the baselines step 13.23 pt inside a
// paragraph, the first is at 80.90 pt and line 48 at 758.88 pt; a line 49 would sit below the
// bottom margin, so Word ends page one there.
async function pageLines(name: string) {
	const { model } = await loadDocx(
		await readFile(
			new URL(`./fixtures/word-layout-reports/issue-383-${name}.docx`, import.meta.url),
		),
	);
	return layoutDocumentModel(model, centuryGothicMeasurer).pages.map((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? block.lines.map((line) => ({
							blockId: block.blockId,
							pt: (page.marginTopPx + block.yPx + line.yPx + (line.baselinePx ?? 0)) * 0.75,
						}))
					: [],
			),
		),
	);
}

it('ends page one after 48 lines at 1.08 line spacing, as Word does (#383)', async () => {
	const pages = await pageLines('line-spacing-259');
	expect(pages.map((lines) => lines.length)).toEqual([48, 49]);
	const [first = [], second = []] = pages;
	expect(Math.abs((first[0]?.pt ?? NaN) - 80.9)).toBeLessThan(0.13);
	expect(Math.abs((first[47]?.pt ?? NaN) - 758.88)).toBeLessThan(0.13);
	expect(Math.abs((first[1]?.pt ?? NaN) - (first[0]?.pt ?? NaN) - 13.23)).toBeLessThan(0.02);
	// Widow control is off: one line of the eighth paragraph stays on page one.
	const last = first.at(-1)?.blockId;
	expect(first.filter((line) => line.blockId === last)).toHaveLength(1);
	expect(second[0]?.blockId).toBe(last);
});

it('moves the eighth paragraph whole to page two with widow control (#383)', async () => {
	const pages = await pageLines('widow-control-on');
	expect(pages.map((lines) => lines.length)).toEqual([47, 48, 2]);
});

it('keeps 52 single-spaced lines on page one (#383 control)', async () => {
	const pages = await pageLines('line-spacing-240-control');
	expect(pages.map((lines) => lines.length)).toEqual([52, 45]);
});
