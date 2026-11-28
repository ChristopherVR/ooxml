import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { createFakeMeasurer, layoutDocumentModel } from './layout/index';
import type { Block, Table } from './model';
import { finalReviewStructure } from './review-final-structure';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const stamp = 'w:author="Author" w:date="2026-01-01T00:00:00Z"';
const page =
	'<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>';

async function docx(body: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr>${page}</w:sectPr></w:body></w:document>`,
	);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const deletedMark = (id: number, inner: string, content: string) =>
	`<w:p><w:pPr><w:rPr><w:del w:id="${id}" ${stamp}/></w:rPr>${inner}</w:pPr>${content}</w:p>`;
const text = (block: Block | undefined): string =>
	block?.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : '';
const grid = '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>';
const deletedRow = (id: number) => `<w:trPr><w:del w:id="${id}" ${stamp}/></w:trPr>`;
const cell = (content: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`;

/** The pages of each laid-out section's text, one string per page. */
function pageTexts(model: Awaited<ReturnType<typeof docx>>): string[] {
	const layout = layoutDocumentModel(model, createFakeMeasurer(), { reviewDisplayMode: 'final' });
	return layout.pages.map(
		(pg) =>
			JSON.stringify(pg)
				.match(/"text":"[^"]*"/g)
				?.map((t) => t.slice(8, -1))
				.join('') ?? '',
	);
}

describe('finalReviewStructure sections', () => {
	it('merges the section whose break paragraph mark was deleted', async () => {
		const model = await docx(
			p('One') +
				deletedMark(
					1,
					`<w:sectPr><w:type w:val="nextPage"/>${page}</w:sectPr>`,
					'<w:r><w:t>Two</w:t></w:r>',
				) +
				p('Three'),
		);
		expect(model.sections).toHaveLength(2);
		const final = finalReviewStructure(model);
		expect(final.blocks.map(text)).toEqual(['One', 'TwoThree']);
		expect(final.sections).toHaveLength(1);
		// No blank first page where the break was.
		expect(pageTexts(final)).toHaveLength(1);
		expect(pageTexts(final)[0]).toContain('One');
	});

	it('merges the section ended by a table that is deleted whole', async () => {
		const model = await docx(
			p('One') +
				`<w:tbl>${grid}<w:tr>${deletedRow(2)}${cell(p('Gone'))}${cell(p('Gone'))}</w:tr></w:tbl>` +
				`<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>${page}</w:sectPr></w:pPr></w:p>` +
				p('Two'),
		);
		const final = finalReviewStructure(model);
		expect(final.blocks.some((block) => block.type === 'table')).toBe(false);
		for (const section of final.sections ?? [])
			expect(final.blocks.map((block) => block.id)).toContain(section.endsAtBlockId);
	});
});

describe('finalReviewStructure tables', () => {
	it('hands a vertical merge on when its first row is deleted', async () => {
		const model = await docx(
			`<w:tbl>${grid}` +
				`<w:tr>${deletedRow(1)}${cell(p('MERGED'), '<w:vMerge w:val="restart"/>')}${cell(p('A'))}</w:tr>` +
				`<w:tr>${cell('<w:p/>', '<w:vMerge/>')}${cell(p('B'))}</w:tr>` +
				`<w:tr>${cell('<w:p/>', '<w:vMerge/>')}${cell(p('C'))}</w:tr></w:tbl>`,
		);
		const final = finalReviewStructure(model);
		const rows = (final.blocks[0] as Table).rows;
		expect(rows).toHaveLength(2);
		expect(rows[0]![0]!.verticalMerge).toBe('restart');
		expect(rows[0]![0]!.paragraphs.map(text).join('')).toBe('MERGED');
		expect(rows[1]![0]!.verticalMerge).toBe('continue');
	});

	it('applies deleted marks and rows inside nested tables', async () => {
		const nested =
			`<w:tbl>${grid}<w:tr>${cell(p('In1'))}${cell(p('In2'))}</w:tr>` +
			`<w:tr>${deletedRow(5)}${cell(p('Dead'))}${cell(p('Dead'))}</w:tr></w:tbl>`;
		const model = await docx(
			`<w:tbl>${grid}<w:tr>${cell(
				deletedMark(3, '', '<w:r><w:t xml:space="preserve">Head, </w:t></w:r>') +
					p('tail') +
					nested +
					p('end'),
			)}${cell(p('Other'))}</w:tr></w:tbl>`,
		);
		const outer = finalReviewStructure(model).blocks[0] as Table;
		const inner = outer.rows[0]![0]!;
		const source = inner.sourceBlocks ?? [];
		expect(source.map((block) => (block.type === 'table' ? 'table' : text(block)))).toEqual([
			'Head, tail',
			'table',
			'end',
		]);
		const nestedTable = source.find((block): block is Table => block.type === 'table');
		expect(nestedTable?.rows).toHaveLength(1);
	});
});
