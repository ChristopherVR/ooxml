import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph } from './model';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import { reviewDocumentFormatting } from './review-document-formatting';
import { documentBlockLists } from './document-paragraphs';

const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraph = (id: string): Paragraph => ({
	type: 'paragraph',
	id,
	align: 'center',
	markRevision: { id: `${id}-mark`, kind: 'insert', author: 'Ada' },
	formatRevision: {
		id: `${id}-p`,
		kind: 'paragraphChange',
		author: 'Ada',
		previousParagraphPropertiesXml: `<w:pPr xmlns:w="${namespace}"><w:jc w:val="right"/></w:pPr>`,
	},
	runs: [
		{
			text: 'Text',
			bold: true,
			formatRevision: {
				id: `${id}-r`,
				kind: 'formatChange',
				author: 'Ada',
				previousRunPropertiesXml: `<w:rPr xmlns:w="${namespace}"><w:i/></w:rPr>`,
			},
		},
		{ text: 'Added', revision: { id: `${id}-i`, kind: 'insert', author: 'Ada' } },
	],
});

describe('document review formatting projection', () => {
	it('projects every story and table cell without changing text offsets, identities or pending text revisions', async () => {
		const model = (await loadDocx(await saveDocx(createDocument()))).model;
		model.blocks = [
			paragraph('body'),
			{ type: 'table', id: 'table', rows: [[{ paragraphs: [paragraph('cell')] }]] },
		];
		const section = model.sections![0]!;
		section.headers = {
			default: { blocks: [paragraph('header')] },
			first: { blocks: [paragraph('first')] },
			even: { blocks: [paragraph('even')] },
		};
		section.footers = { default: { blocks: [paragraph('footer')] } };
		model.footnotes = [{ id: '2', blocks: [paragraph('footnote')] }];
		model.endnotes = [{ id: '3', blocks: [paragraph('endnote')] }];
		const before = structuredClone(model);
		const projected = reviewDocumentFormatting(model, 'original');
		expect(projected.diagnostics).toEqual([]);
		const paragraphs = documentBlockLists(projected.model).flatMap((blocks) =>
			blocks.flatMap((block) =>
				block.type === 'paragraph'
					? [block]
					: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
			),
		);
		expect(paragraphs).toHaveLength(8);
		for (const value of paragraphs) {
			expect(value.align).toBe('right');
			expect(value.runs[0]).toMatchObject({ text: 'Text', italic: true });
			expect(value.runs[0]!.bold).toBeUndefined();
			expect(value.runs[1]!.revision!.id).toBe(`${value.id}-i`);
			expect(value.markRevision!.id).toBe(`${value.id}-mark`);
			expect(value.runs.map((run) => run.text).join('')).toBe('TextAdded');
		}
		expect(model).toEqual(before);
		for (const mode of ['all', 'simple', 'final'] as const)
			expect(reviewDocumentFormatting(model, mode).model).toBe(model);
	});

	it('reports individual unavailable snapshots while projecting other valid histories', () => {
		const model = createDocument();
		const invalid = paragraph('invalid');
		delete invalid.formatRevision!.previousParagraphPropertiesXml;
		delete invalid.runs[0]!.formatRevision!.previousRunPropertiesXml;
		model.blocks = [invalid, paragraph('valid')];
		const before = structuredClone(model);
		const projected = reviewDocumentFormatting(model, 'original');
		expect(projected.diagnostics).toEqual([
			{ paragraphId: 'invalid', message: expect.any(String) },
			{ paragraphId: 'invalid', runIndex: 0, message: expect.any(String) },
		]);
		expect(projected.model.blocks[0]).toEqual(invalid);
		expect(projected.model.blocks[1]).toMatchObject({
			align: 'right',
			runs: [{ text: 'Text', italic: true }, { text: 'Added' }],
		});
		expect(model).toEqual(before);
	});
});
