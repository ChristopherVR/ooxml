import { describe, expect, it } from 'vitest';
import { cellSourceBlocks } from './cell-blocks';
import { loadDocx, type Block, type DocumentModel, type Paragraph } from './index';
import { refreshStyleRefFields } from './style-ref-fields';
import { at, expectParagraph, must } from './test-support/access';
import {
	complexField,
	paragraph,
	simpleField,
	styleRefFixture,
} from './test-support/style-ref-fixture';

const load = async (parts: Parameters<typeof styleRefFixture>[0]) =>
	(await loadDocx(await styleRefFixture(parts))).model;
const headerParagraph = (model: DocumentModel): Paragraph =>
	expectParagraph(must(at(model.sections, 0).headers?.default, 'header').blocks[0]);
const shown = (item: Paragraph) => item.runs.map((run) => run.text).join('');

describe('refreshing STYLEREF results', () => {
	it('gives a header field the text it shows on every page, in its first result run', async () => {
		const model = await load({
			body: paragraph('DOC-0042', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				complexField('STYLEREF "Doc Number" \\* MERGEFORMAT', ['DOC', '-1']),
			),
			footer: paragraph('', undefined, simpleField('STYLEREF DocNumber', 'DOC-1')),
		});
		const refreshed = refreshStyleRefFields(model);
		const header = headerParagraph(refreshed);
		const results = header.runs.filter((run) => run.field);
		expect(results.map((run) => run.text)).toEqual(['DOC-0042', '']);
		expect(results[0]!.bold).toBe(true);
		const footer = expectParagraph(
			must(at(refreshed.sections, 0).footers?.default, 'footer').blocks[0],
		);
		expect(shown(footer)).toBe('DOC-0042');
		expect(refreshed.blocks).toBe(model.blocks);
	});

	it('formats a CHARFORMAT result like the first letter of the field name', async () => {
		const field = (format: string) =>
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
			'<w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>' +
			`<w:r><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr><w:instrText xml:space="preserve">STYLEREF "Doc Number" \\* ${format} </w:instrText></w:r>` +
			'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
			'<w:r><w:rPr><w:b/></w:rPr><w:t>DOC-1</w:t></w:r>' +
			'<w:r><w:fldChar w:fldCharType="end"/></w:r>';
		const model = await load({
			body: paragraph('DOC-2', 'DocNumber'),
			header: paragraph('', undefined, field('CHARFORMAT') + field('MERGEFORMAT')),
		});
		const results = headerParagraph(refreshStyleRefFields(model)).runs.filter((run) => run.field);
		expect(results.map((run) => run.text)).toEqual(['DOC-2', 'DOC-2']);
		expect(results.map((run) => [run.bold, run.italic, run.color])).toEqual([
			[undefined, true, '#C00000'],
			[true, undefined, undefined],
		]);
		// The formatting is brought up to date when the text already is; MERGEFORMAT keeps its own.
		const current = await load({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph('', undefined, field('CHARFORMAT') + field('MERGEFORMAT')),
		});
		const formatted = headerParagraph(refreshStyleRefFields(current)).runs.filter(
			(run) => run.field,
		);
		expect(formatted.map((run) => [run.text, run.bold, run.italic])).toEqual([
			['DOC-1', undefined, true],
			['DOC-1', true, undefined],
		]);
		const merged = await load({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph('', undefined, field('MERGEFORMAT')),
		});
		expect(refreshStyleRefFields(merged)).toBe(merged);
		expect(refreshStyleRefFields(refreshStyleRefFields(current))).toEqual(
			refreshStyleRefFields(current),
		);
	});

	it('returns the model itself when every result is current', async () => {
		const model = await load({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1'])),
		});
		expect(refreshStyleRefFields(model)).toBe(model);
	});

	it('keeps a header result that depends on the page, a locked field and an unknown style', async () => {
		const model = await load({
			body: paragraph('One', 'Heading1') + paragraph('Two', 'Heading1'),
			header: paragraph(
				'',
				undefined,
				complexField('STYLEREF 1', ['Cached']) +
					complexField(
						'STYLEREF "Doc Number"',
						['Locked'],
						'<w:fldChar w:fldCharType="begin" w:fldLock="1"/>',
					) +
					complexField('STYLEREF Missing', ['Kept']),
			),
		});
		const withNumber = {
			...model,
			blocks: [
				...model.blocks,
				{ type: 'paragraph' as const, id: 'doc', style: 'DocNumber', runs: [{ text: 'DOC-9' }] },
			],
		};
		expect(refreshStyleRefFields(withNumber).sections).toBe(withNumber.sections);
	});

	it('updates a body field from the nearest text before it, else after it, only when asked', async () => {
		const model = await load({
			body:
				paragraph('Scope', 'Heading1') +
				paragraph('', undefined, complexField('STYLEREF 1', ['old'])) +
				paragraph('Method', 'Heading1') +
				paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['old'])) +
				paragraph('DOC-3', 'DocNumber'),
		});
		expect(refreshStyleRefFields(model)).toBe(model);
		const refreshed = refreshStyleRefFields(model, { body: true });
		expect(shown(expectParagraph(refreshed.blocks[1]))).toBe('Scope');
		expect(shown(expectParagraph(refreshed.blocks[3]))).toBe('DOC-3');
		expect(refreshed.blocks[0]).toBe(model.blocks[0]);
	});

	it('searches a body field own paragraph first, back from the field and then on from it', async () => {
		const styled = (text: string) =>
			`<w:r><w:rPr><w:rStyle w:val="EntryName"/></w:rPr><w:t>${text}</w:t></w:r>`;
		const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
		const field = complexField('STYLEREF "Entry Name"', ['old']);
		// A result in the style it refers to is the field's own, not text it can take.
		const own =
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
			'<w:r><w:instrText xml:space="preserve"> STYLEREF "Entry Name" </w:instrText></w:r>' +
			`<w:r><w:fldChar w:fldCharType="separate"/></w:r>${styled('stale')}` +
			'<w:r><w:fldChar w:fldCharType="end"/></w:r>';
		const model = await load({
			body:
				paragraph('', undefined, `${field}${plain(' then ')}${styled('Oak')}`) +
				paragraph('', undefined, styled('Alder')) +
				paragraph('', undefined, `${styled('Ash')}${plain(', see ')}${field}`) +
				paragraph('', undefined, styled('Pine')) +
				paragraph('', undefined, own),
		});
		const refreshed = refreshStyleRefFields(model, { body: true });
		expect(shown(expectParagraph(refreshed.blocks[0]))).toBe('Oak then Oak');
		expect(shown(expectParagraph(refreshed.blocks[2]))).toBe('Ash, see Ash');
		expect(shown(expectParagraph(refreshed.blocks[4]))).toBe('Pine');
	});

	it('updates fields in tables nested in a cell, in headers and in the body', async () => {
		const cell = (inner: string) =>
			`<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
		const table = (inner: string) =>
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
		const nested = (instr: string) =>
			table(
				cell(
					paragraph('Label') +
						table(cell(paragraph('', undefined, complexField(instr, ['old'])))) +
						paragraph(''),
				),
			);
		const model = await load({
			body: paragraph('DOC-7', 'DocNumber') + nested('STYLEREF "Doc Number"'),
			header: nested('STYLEREF "Doc Number" \\* MERGEFORMAT'),
		});
		const fieldText = (blocks: readonly Block[]) => {
			const found: string[] = [];
			const visit = (items: readonly Block[]) => {
				for (const block of items)
					if (block.type === 'paragraph') {
						if (block.runs.some((run) => run.field)) found.push(shown(block));
					} else
						for (const row of block.rows) for (const item of row) visit(cellSourceBlocks(item));
			};
			visit(blocks);
			return found;
		};
		const header = must(at(model.sections, 0).headers?.default, 'header');
		const refreshed = refreshStyleRefFields(model);
		expect(fieldText(must(at(refreshed.sections, 0).headers?.default, 'header').blocks)).toEqual([
			'DOC-7',
		]);
		expect(fieldText(header.blocks)).toEqual(['old']);
		expect(fieldText(refreshed.blocks)).toEqual(['old']);
		expect(fieldText(refreshStyleRefFields(model, { body: true }).blocks)).toEqual(['DOC-7']);
	});

	it('writes a result into the text a revision keeps, never into deleted text', async () => {
		const revised = (deleted: string, inserted?: string) =>
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
			'<w:r><w:instrText xml:space="preserve"> STYLEREF "Doc Number" </w:instrText></w:r>' +
			'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
			`<w:del w:id="1" w:author="A"><w:r><w:delText>${deleted}</w:delText></w:r></w:del>` +
			(inserted === undefined
				? ''
				: `<w:ins w:id="2" w:author="A"><w:r><w:t>${inserted}</w:t></w:r></w:ins>`) +
			'<w:r><w:fldChar w:fldCharType="end"/></w:r>';
		const results = (model: DocumentModel) =>
			headerParagraph(model)
				.runs.filter((run) => run.field)
				.map((run) => [run.text, run.revision?.kind]);
		const replaced = await load({
			body: paragraph('DOC-3', 'DocNumber'),
			header: paragraph('', undefined, revised('DOC-1', 'DOC-2')),
		});
		expect(results(refreshStyleRefFields(replaced))).toEqual([
			['DOC-1', 'delete'],
			['DOC-3', 'insert'],
		]);
		// The text the revision keeps is current: nothing changes.
		const current = await load({
			body: paragraph('DOC-2', 'DocNumber'),
			header: paragraph('', undefined, revised('DOC-1', 'DOC-2')),
		});
		expect(refreshStyleRefFields(current)).toBe(current);
		// A result a revision deletes whole stays as it is until the revision is settled.
		const deleted = await load({
			body: paragraph('DOC-3', 'DocNumber'),
			header: paragraph('', undefined, revised('DOC-1')),
		});
		expect(refreshStyleRefFields(deleted)).toBe(deleted);
		// A field a revision deletes keeps its result.
		const gone = await load({
			body: paragraph('DOC-3', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				'<w:del w:id="3" w:author="A"><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
					'<w:r><w:delInstrText xml:space="preserve"> STYLEREF "Doc Number" </w:delInstrText></w:r>' +
					'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>DOC-1</w:delText></w:r>' +
					'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>',
			),
		});
		expect(refreshStyleRefFields(gone)).toBe(gone);
	});

	it('updates a field whose switches come before its style', async () => {
		const model = await load({
			body: paragraph('DOC-9', 'DocNumber'),
			header: paragraph('', undefined, complexField('STYLEREF \\* Upper "Doc Number"', ['doc-1'])),
		});
		expect(shown(headerParagraph(refreshStyleRefFields(model)))).toBe('DOC-9');
	});

	it("gives a numbered paragraph's own number to a STYLEREF \\n in it", async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const item = (result: string) =>
			paragraph(
				'',
				'ListNumber',
				`<w:r><w:t xml:space="preserve">Deny paragraph </w:t></w:r>${complexField('STYLEREF "List Number" \\n', [result])}`,
			);
		const model = await load({
			styles: `<w:styles xmlns:w="${w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>`,
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
			body:
				item('1') +
				item('2') +
				item('9') +
				paragraph('', undefined, complexField('STYLEREF "List Number" \\n', ['old'])),
		});
		const results = (current: DocumentModel) =>
			current.blocks.map((block) =>
				expectParagraph(block)
					.runs.filter((run) => run.field)
					.map((run) => run.text)
					.join(''),
			);
		// Its own paragraph is the nearest in the style, as when Word repeats a paragraph's number
		// in its text; a paragraph out of the style reads the one before it.
		expect(results(refreshStyleRefFields(model, { body: true }))).toEqual(['1', '2', '3', '3']);
	});
});
