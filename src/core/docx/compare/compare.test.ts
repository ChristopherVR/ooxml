import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { Block, DocumentModel } from '../model';
import { loadDocx } from '../parse';
import { acceptAllRevisions, listRevisions, rejectAllRevisions } from '../revision-commands';
import { acceptForComparison, compareDocuments } from './compare';
import { revisionSpans } from './spans';
import { UNCOMPARED_DOCUMENT_PARTS } from './types';

const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function docx(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const p = (...texts: string[]) =>
	texts.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('');
const tbl = (...rows: string[][]) =>
	`<w:tbl><w:tblPr/><w:tblGrid>${rows[0]!.map(() => '<w:gridCol w:w="2000"/>').join('')}</w:tblGrid>${rows
		.map((cells) => `<w:tr>${cells.map((cell) => `<w:tc>${p(cell)}</w:tc>`).join('')}</w:tr>`)
		.join('')}</w:tbl>`;
const OPTIONS = { author: 'Compare Author', date: '2026-01-01T00:00:00Z' };

/** Compares, saves, and checks that rejecting gives `before` and accepting gives `after`. */
async function roundTrip(original: string, revised: string, before: string[], after: string[]) {
	const result = await compareDocuments(await docx(original), await docx(revised), OPTIONS);
	expect(bodyText(rejectAllRevisions(result.model))).toEqual(before);
	expect(bodyText(acceptAllRevisions(result.model))).toEqual(after);
	return result;
}

/** Paragraph texts of a story, tables flattened cell by cell. */
function texts(blocks: Block[]): string[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block.runs.map((run) => run.text).join('')]
			: block.rows.flatMap((row) => row.flatMap((cell) => texts(cell.paragraphs))),
	);
}
const bodyText = (model: DocumentModel) => texts(model.blocks);

describe('compareDocuments', () => {
	it('records nothing for identical documents and lists what it never compares', async () => {
		const bytes = await docx(p('Same text.', 'Second.'));
		const result = await compareDocuments(bytes, bytes, OPTIONS);
		expect(result.summary).toEqual({ insertions: 0, deletions: 0, formatChanges: 0 });
		expect(listRevisions(result.model)).toEqual([]);
		expect(result.notCompared).toEqual(UNCOMPARED_DOCUMENT_PARTS);
	});

	it('marks changed words, stamps the author and date and numbers ids as decimals', async () => {
		const result = await compareDocuments(
			await docx(p('One two three.')),
			await docx(p('One 2 three.')),
			OPTIONS,
		);
		expect(revisionSpans(result.model)).toEqual([
			{ kind: 'unchanged', text: 'One ' },
			{ kind: 'delete', text: 'two' },
			{ kind: 'insert', text: '2' },
			{ kind: 'unchanged', text: ' three.¶' },
		]);
		const revisions = listRevisions(result.model);
		expect(revisions.map((entry) => entry.author)).toEqual(['Compare Author', 'Compare Author']);
		expect(revisions.every((entry) => entry.date === '2026-01-01T00:00:00Z')).toBe(true);
		expect(revisions.every((entry) => /^\d+$/.test(entry.id))).toBe(true);
		expect(new Set(revisions.map((entry) => entry.id)).size).toBe(2);
		expect(result.summary).toEqual({ insertions: 1, deletions: 1, formatChanges: 0 });
		const xml = await (
			await JSZip.loadAsync(result.bytes)
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('<w:delText>two</w:delText>');
	});

	it('marks only the changed letters at character granularity', async () => {
		const result = await compareDocuments(await docx(p('colour')), await docx(p('color')), {
			...OPTIONS,
			granularity: 'character',
		});
		expect(revisionSpans(result.model)).toEqual([
			{ kind: 'unchanged', text: 'colo' },
			{ kind: 'delete', text: 'u' },
			{ kind: 'unchanged', text: 'r¶' },
		]);
	});

	it('shows a paragraph added at the end by marking the previous mark, as Word does', async () => {
		const result = await compareDocuments(
			await docx(p('Kept.')),
			await docx(p('Kept.', 'Added.')),
			OPTIONS,
		);
		expect(revisionSpans(result.model)).toEqual([
			{ kind: 'unchanged', text: 'Kept.' },
			{ kind: 'insert', text: '¶Added.' },
			{ kind: 'unchanged', text: '¶' },
		]);
	});

	it('shows a deleted middle paragraph with its own mark', async () => {
		const result = await compareDocuments(
			await docx(p('First.', 'Gone.', 'Last.')),
			await docx(p('First.', 'Last.')),
			OPTIONS,
		);
		expect(revisionSpans(result.model)).toEqual([
			{ kind: 'unchanged', text: 'First.¶' },
			{ kind: 'delete', text: 'Gone.¶' },
			{ kind: 'unchanged', text: 'Last.¶' },
		]);
	});

	it('records a formatting change of unchanged text with the original properties', async () => {
		const result = await compareDocuments(
			await docx('<w:p><w:r><w:t>Plain</w:t></w:r></w:p>'),
			await docx('<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Plain</w:t></w:r></w:p>'),
			OPTIONS,
		);
		const [entry] = listRevisions(result.model);
		expect(entry).toMatchObject({ kind: 'formatChange', author: 'Compare Author' });
		const xml = await (
			await JSZip.loadAsync(result.bytes)
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toMatch(/<w:b\/><w:i\/><w:rPrChange [^>]*><w:rPr\/><\/w:rPrChange>/);
		const off = await compareDocuments(
			await docx('<w:p><w:r><w:t>Plain</w:t></w:r></w:p>'),
			await docx('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Plain</w:t></w:r></w:p>'),
			{ ...OPTIONS, compareFormatting: false },
		);
		expect(listRevisions(off.model)).toEqual([]);
	});

	it('compares field codes whole, never splitting them', async () => {
		const field = (instr: string) =>
			`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
		const result = await compareDocuments(
			await docx(`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field(' PAGE ')}</w:p>`),
			await docx(
				`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field(' NUMPAGES ')}</w:p>`,
			),
			OPTIONS,
		);
		expect(revisionSpans(result.model)).toEqual([{ kind: 'unchanged', text: 'Page 1¶' }]);
		const paragraph = result.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph');
		const codes = paragraph.runs.filter((run) => run.fieldCode !== undefined);
		expect(codes.map((run) => [run.fieldCode?.trim(), run.revision?.kind])).toEqual([
			['PAGE', 'delete'],
			['NUMPAGES', 'insert'],
		]);
	});

	it('accepts the tracked changes of both documents before comparing', async () => {
		const tracked = await docx(
			'<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:ins w:id="1" w:author="X"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2" w:author="X"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>',
		);
		const result = await compareDocuments(tracked, await docx(p('A new')), OPTIONS);
		expect(listRevisions(result.model)).toEqual([]);
	});

	it('keeps a tracked mark that cannot be accepted instead of failing', () => {
		const model: DocumentModel = {
			blocks: [
				{
					type: 'paragraph',
					id: 'last',
					runs: [{ text: 'End' }],
					markRevision: { id: '1', kind: 'delete', author: 'X' },
				},
			],
			page: {
				width: 816,
				height: 1056,
				marginTop: 0,
				marginRight: 0,
				marginBottom: 0,
				marginLeft: 0,
			},
			warnings: [],
		};
		const warnings: string[] = [];
		const accepted = acceptForComparison(model, warnings);
		expect(listRevisions(accepted)).toEqual([]);
		expect(bodyText(accepted)).toEqual(['End']);
		expect(warnings).toHaveLength(1);
	});

	it('keeps an inserted table above the paragraph that follows it', async () => {
		const result = await roundTrip(
			p('Kept.', 'Old'),
			p('Kept.') + tbl(['T']) + p('New'),
			['Kept.', 'Old'],
			['Kept.', 'T', 'New'],
		);
		expect(result.model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'table',
			'paragraph',
		]);
	});

	it('keeps text added above a changed table above it', async () => {
		const result = await roundTrip(
			p('Kept.') + tbl(['A', 'one']) + p(''),
			p('Kept.', 'Added.') + tbl(['A', 'two']) + p(''),
			['Kept.', 'A', 'one', ''],
			['Kept.', 'Added.', 'A', 'two', ''],
		);
		expect(result.model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'paragraph',
			'table',
			'paragraph',
		]);
	});

	it('shows a table with a column added as replaced, so it still saves', async () => {
		const result = await roundTrip(
			p('Kept.') + tbl(['A', 'B']) + p(''),
			p('Kept.') + tbl(['A', 'B', 'C']) + p(''),
			['Kept.', 'A', 'B', ''],
			['Kept.', 'A', 'B', 'C', ''],
		);
		expect(result.warnings.join(' ')).toContain('deleted and inserted whole');
	});

	it('shows a paragraph removed or added just above a table', async () => {
		await roundTrip(
			p('Kept.', 'Gone.') + tbl(['T']) + p(''),
			p('Kept.') + tbl(['T']) + p(''),
			['Kept.', 'Gone.', 'T', ''],
			['Kept.', 'T', ''],
		);
		await roundTrip(
			p('Kept.') + tbl(['T']) + p(''),
			p('Kept.', 'Added.') + tbl(['T']) + p(''),
			['Kept.', 'T', ''],
			['Kept.', 'Added.', 'T', ''],
		);
	});

	it('leaves tables with nested tables alone when they did not change', async () => {
		const bytes = await readFile(
			join(__dirname, '..', '__fixtures__', 'nested-three-row', 'nested-three-short.docx'),
		);
		const result = await compareDocuments(bytes, bytes, OPTIONS);
		expect(result.summary).toEqual({ insertions: 0, deletions: 0, formatChanges: 0 });
	});

	it('shows a paragraph whose runs the writer cannot split inserted whole', async () => {
		const corpus = join(__dirname, '..', '__fixtures__', 'word-corpus');
		const [original, revised] = await Promise.all([
			readFile(join(corpus, '13-equation-shapes.docx')),
			readFile(join(corpus, '09-fields-toc.docx')),
		]);
		const result = await compareDocuments(original, revised, OPTIONS);
		expect(result.warnings.join(' ')).toContain('shown inserted whole');
		const warnings: string[] = [];
		expect(bodyText(acceptAllRevisions(result.model))).toEqual(
			bodyText(acceptForComparison((await loadDocx(revised)).model, warnings)),
		);
	});

	it('rejecting every change gives the original back and accepting gives the revised', async () => {
		const fixtures = join(__dirname, '..', '__fixtures__', 'compare');
		for (const name of ['words', 'format', 'tracked', 'chars', 'table']) {
			const [original, revised] = await Promise.all([
				readFile(join(fixtures, `${name}-original.docx`)),
				readFile(join(fixtures, `${name}-revised.docx`)),
			]);
			const result = await compareDocuments(original, revised, {
				...OPTIONS,
				granularity: name === 'chars' ? 'character' : 'word',
			});
			const warnings: string[] = [];
			const expectedOriginal = acceptForComparison((await loadDocx(original)).model, warnings);
			const expectedRevised = acceptForComparison((await loadDocx(revised)).model, warnings);
			expect(bodyText(rejectAllRevisions(result.model)), name).toEqual(bodyText(expectedOriginal));
			expect(bodyText(acceptAllRevisions(result.model)), name).toEqual(bodyText(expectedRevised));
		}
	});
});
