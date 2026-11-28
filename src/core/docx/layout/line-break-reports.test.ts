import { readFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createEstimatedMeasurer, createFontFileMeasurer } from '../pdf-export/font-measurer';
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

describe('a content control start tag after a word is no break opportunity (#410)', () => {
	it('moves the word down with the tags and their text, justified and left', async () => {
		for (const paragraph of await lines('issue-410-start-tag-after-word', estimated)) {
			expect(paragraph[0]?.endsWith('amet')).toBe(true);
			expect(paragraph[1]?.startsWith('Name/Field')).toBe(true);
			expect(paragraph[1]?.endsWith('iota')).toBe(true);
		}
	});

	it('breaks at the space when one separates the word and the tags', async () => {
		for (const paragraph of await lines('issue-410-start-tag-after-space', estimated)) {
			expect(paragraph[0]?.endsWith('Name/Field')).toBe(true);
			expect(paragraph[1]?.startsWith('Outer')).toBe(true);
		}
	});
});

describe('small caps draw spaces at the reduced size (#392)', () => {
	const files = createFontFileMeasurer([
		{
			family: 'Liberation Serif',
			data: new Uint8Array(
				readFileSync(
					new URL(
						'../pdf-export/__fixtures__/fonts/LiberationSerif-Regular.subset.ttf',
						import.meta.url,
					),
				),
			),
		},
	]);
	const spec = { family: 'Times New Roman', sizePx: (11 * 96) / 72 };

	it('measures a space at 0.8 of the size, and capitals at full size', () => {
		const plain = files.widthOf(' ', spec);
		expect(files.widthOf(' ', { ...spec, smallCaps: true })).toBeCloseTo(plain * 0.8, 6);
		expect(files.widthOf('AB', { ...spec, smallCaps: true })).toBe(files.widthOf('AB', spec));
		expect(files.widthOf('AB AB', { ...spec, smallCaps: true })).toBeCloseTo(
			files.widthOf('AB', spec) * 2 + plain * 0.8,
			6,
		);
	});

	it('puts 26 AB words on line 1, 25 without small caps, as Word does', async () => {
		const small = await lines('issue-392-small-caps-spaces', files);
		expect(
			small[0]?.[0]
				?.trimEnd()
				.split(' ')
				.filter((word) => word === 'AB'),
		).toHaveLength(26);
		expect(small[0]?.[1]?.trim()).toBe('NEXT LINE');
		const control = await lines('issue-392-small-caps-spaces-control', files);
		expect(
			control[0]?.[0]
				?.trimEnd()
				.split(' ')
				.filter((word) => word === 'AB'),
		).toHaveLength(25);
	});
});
