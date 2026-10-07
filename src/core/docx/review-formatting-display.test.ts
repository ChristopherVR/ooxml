import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './parse';
import { reviewParagraphFormatting, reviewRunFormatting } from './review-formatting-display';
import { expectParagraph, paragraphWithoutXmlBases } from './test-support/access';
import type { TextRun } from './model';

describe('review formatting display projections', () => {
	for (const name of ['alignment', 'spacing', 'indent', 'multiple'])
		it(`matches native rejected ${name} paragraph properties without altering its history`, async () => {
			const load = async (suffix: string) =>
				loadDocx(
					new Uint8Array(
						await readFile(
							new URL(
								`./__fixtures__/review-paragraph-formatting/${name}-${suffix}.docx`,
								import.meta.url,
							),
						),
					),
				);
			const current = expectParagraph((await load('tracked')).model.blocks[0]);
			const before = structuredClone(current);
			const expected = expectParagraph((await load('rejected')).model.blocks[0]);
			const projected = reviewParagraphFormatting(current, 'original');
			expect(projected.error).toBeUndefined();
			expect(paragraphWithoutXmlBases(projected.value)).toEqual(paragraphWithoutXmlBases(expected));
			expect(projected.value.runs).toBe(current.runs);
			expect(current).toEqual(before);
			for (const mode of ['all', 'simple', 'final'] as const)
				expect(reviewParagraphFormatting(current, mode).value).toBe(current);
		});

	for (const name of ['bold', 'multiple'])
		it(`matches native rejected ${name} run properties without altering text or history`, async () => {
			const load = async (suffix: string) =>
				loadDocx(
					new Uint8Array(
						await readFile(
							new URL(`./__fixtures__/review-formatting/${name}-${suffix}.docx`, import.meta.url),
						),
					),
				);
			const current = expectParagraph((await load('tracked')).model.blocks[0]);
			const before = structuredClone(current);
			const expected = expectParagraph((await load('rejected')).model.blocks[0]);
			const projected = current.runs.map((run) => reviewRunFormatting(run, 'original'));
			expect(projected.every((run) => run.error === undefined)).toBe(true);
			expect(projected.map(({ value: { restoredRunPropertiesXml: _xml, ...run } }) => run)).toEqual(
				expected.runs,
			);
			expect(current).toEqual(before);
		});

	it('restores run properties while keeping an overlapping text revision and anchors', () => {
		const run: TextRun = {
			text: 'Text',
			bold: true,
			revision: { id: 'insert', kind: 'insert', author: 'Bob' },
			formatRevision: {
				id: 'format',
				kind: 'formatChange',
				author: 'Ada',
				previousRunPropertiesXml:
					'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:i/></w:rPr>',
			},
		};
		const before = structuredClone(run);
		const projected = reviewRunFormatting(run, 'original');
		expect(projected.value).toMatchObject({ text: 'Text', italic: true, revision: run.revision });
		expect(projected.value.bold).toBeUndefined();
		expect(projected.value.formatRevision).toBeUndefined();
		expect(run).toEqual(before);
		expect(reviewRunFormatting(run, 'final').value).toBe(run);
	});

	it('reports unavailable history and returns the untouched source', () => {
		for (const previousRunPropertiesXml of [undefined, '<rPr xmlns="urn:wrong"/>']) {
			const run: TextRun = {
				text: 'Text',
				bold: true,
				revision: {
					id: '7',
					kind: 'formatChange',
					author: 'Ada',
					...(previousRunPropertiesXml ? { previousRunPropertiesXml } : {}),
				},
			};
			const before = structuredClone(run);
			const projected = reviewRunFormatting(run, 'original');
			expect(projected.value).toBe(run);
			expect(projected.error).toBeTruthy();
			expect(run).toEqual(before);
		}
	});
});
