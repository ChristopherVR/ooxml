import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createEstimatedMeasurer } from '../pdf-export/font-measurer';
import { layoutDocumentModel } from './index';
import type { TextMeasurer } from './measure';

// Line-breaking reports measured against Word 16 (build 16.0.20625, compatibility mode 15). The
// fixtures are rebuilt from the paragraphs quoted in each issue by
// `fixtures/word-layout-reports/generate-line-break-fixtures.mjs`; the Word results asserted here
// are the ones the issues give.

/** The text of each line of each paragraph, in document order. */
async function lines(name: string, measurer: TextMeasurer): Promise<string[][]> {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, measurer);
	const seen = new Set<string>();
	const out: string[][] = [];
	for (const page of pages)
		for (const column of page.columns)
			for (const block of column.blocks) {
				if (block.kind !== 'paragraph' || seen.has(block.blockId)) continue;
				seen.add(block.blockId);
				out.push(block.lines.map((line) => line.fragments.map((f) => f.text).join('')));
			}
	return out;
}

const estimated = createEstimatedMeasurer();
const endsWith = (paragraph: string[] | undefined, word: string) =>
	paragraph?.[0]?.trimEnd().endsWith(word);

describe('a space next to a non-breaking space is no break opportunity (#406)', () => {
	it('moves John down with SMITH, in both orders, as Word does', async () => {
		const bug = await lines('issue-406-space-next-to-nbsp', estimated);
		expect(bug).toHaveLength(12);
		// Paragraphs 1 and 2 fit "John SMITH,"; the other ten end line 1 at "Dr" in Word.
		const first = bug.map((paragraph) => (endsWith(paragraph, 'SMITH,') ? 'fits' : 'Dr'));
		expect(first).toEqual([
			'fits',
			'fits',
			...Array(4).fill('Dr'),
			'fits',
			'fits',
			...Array(4).fill('Dr'),
		]);
		for (const [index, paragraph] of bug.entries())
			if (first[index] === 'Dr') expect(paragraph[1]?.startsWith('John')).toBe(true);
	});

	it('still breaks between two ordinary spaces', async () => {
		const control = await lines('issue-406-two-spaces-control', estimated);
		expect(control.map((paragraph) => paragraph[0]?.trimEnd().split(' ').at(-1))).toEqual([
			'SMITH,',
			'SMITH,',
			'John',
			'John',
			'John',
			'Dr',
		]);
	});
});
