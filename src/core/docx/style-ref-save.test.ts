import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { DocxModelValidationError, loadDocx, type DocumentModel } from './index';
import { expectParagraph, must } from './test-support/access';
import {
	complexField,
	paragraph,
	simpleField,
	styleRefFixture,
} from './test-support/style-ref-fixture';

const HEADER = paragraph(
	'',
	undefined,
	`<w:r><w:t xml:space="preserve">Document </w:t></w:r>${complexField('STYLEREF "Doc Number" \\* MERGEFORMAT', ['DOC', '-1'])}`,
);
const FOOTER = paragraph('', undefined, simpleField('STYLEREF Title', 'Plan'));

async function parts(bytes: Uint8Array): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const found = new Map<string, string>();
	for (const file of Object.values(zip.files))
		if (!file.dir) found.set(file.name, await file.async('string'));
	return found;
}

/** Sets the text of the body paragraph in `style`. */
function setText(model: DocumentModel, style: string, text: string): DocumentModel {
	const next = structuredClone(model);
	const target = next.blocks.find((block) => block.type === 'paragraph' && block.style === style);
	expectParagraph(target).runs = [{ text }];
	return next;
}

describe('saving STYLEREF fields', () => {
	it('refreshes header and footer results from the edited body and leaves other parts alone', async () => {
		const original = await styleRefFixture({
			body: paragraph('Plan', 'Title') + paragraph('DOC-1', 'DocNumber') + paragraph('Notes'),
			header: HEADER,
			footer: FOOTER,
		});
		const loaded = await loadDocx(original);
		const saved = await loaded.save(
			setText(setText(loaded.model, 'DocNumber', 'DOC-0042'), 'Title', 'Survey'),
		);
		const before = await parts(original);
		const after = await parts(saved);
		const header = after.get('word/header1.xml')!;
		expect(header).toContain('>Document <');
		expect(header).toMatch(/<w:b\/><\/w:rPr><w:t[^>]*>DOC-0042<\/w:t>/);
		expect(header).not.toContain('-1<');
		expect(after.get('word/footer1.xml')).toContain('>Survey<');
		for (const name of ['word/styles.xml', '[Content_Types].xml', 'word/_rels/document.xml.rels'])
			expect(after.get(name), name).toBe(before.get(name));
		// Reopened, the header still holds one field whose result is the new number.
		const reopened = (await loadDocx(saved)).model;
		const runs = expectParagraph(reopened.sections![0]!.headers!.default!.blocks[0]).runs;
		expect(
			runs
				.filter((run) => run.field)
				.map((run) => run.text)
				.join(''),
		).toBe('DOC-0042');
		expect(runs.filter((run) => run.fieldChar)).toHaveLength(3);
	});

	it('returns the original bytes for an unchanged document, even with a stale header result', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-0042', 'DocNumber'),
			header: HEADER,
		});
		const loaded = await loadDocx(original);
		expect(await loaded.save()).toEqual(original);
	});

	it('does not rewrite a header whose result is already current', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber') + paragraph('Notes'),
			header: paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1'])),
		});
		const loaded = await loadDocx(original);
		const model = structuredClone(loaded.model);
		expectParagraph(model.blocks[1]).runs = [{ text: 'More notes' }];
		const saved = await parts(await loaded.save(model));
		expect(saved.get('word/header1.xml')).toBe((await parts(original)).get('word/header1.xml'));
	});

	it('refreshes a header field in a table nested in a cell', async () => {
		const cell = (inner: string) =>
			`<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
		const table = (inner: string) =>
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: table(
				cell(
					paragraph('Number') +
						table(
							cell(paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1']))),
						) +
						// A field in the cell's own paragraph, beside the nested table.
						paragraph('', undefined, complexField('STYLEREF "Doc Number"', ['DOC-1'])),
				),
			),
		});
		const loaded = await loadDocx(original);
		const saved = await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-0042'));
		const header = (await parts(saved)).get('word/header1.xml')!;
		expect(header).toMatch(
			/<w:tbl>.*<w:tbl>.*>DOC-0042<\/w:t>.*<\/w:tbl>.*>DOC-0042<\/w:t>.*<\/w:tbl>/s,
		);
		expect(header).not.toContain('>DOC-1<');
		expect(header).toContain('>Number<');
	});

	it('reports an invalid model with its typed error, not one from the refresh', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const original = await styleRefFixture({
			body: paragraph(
				'',
				undefined,
				'<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Scope</w:t></w:r>',
			),
			header: paragraph('', undefined, complexField('STYLEREF 1 \\s', ['0'])),
			numbering: `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Chapter %1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
		});
		const loaded = await loadDocx(original);
		const model = structuredClone(loaded.model);
		const level = must(model.numberingCatalog?.abstractNums['0']?.levels[0], 'level');
		level.lvlText = 7 as never;
		const error = await loaded.save(model).then(
			() => undefined,
			(caught: unknown) => caught,
		);
		expect(error).toBeInstanceOf(DocxModelValidationError);
		expect((error as Error).message).toContain('lvlText');
	});

	it('writes a CHARFORMAT header result in the formatting of its field code', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
					'<w:r><w:rPr><w:i/></w:rPr><w:instrText xml:space="preserve"> STYLEREF "Doc Number" \\* CHARFORMAT </w:instrText></w:r>' +
					'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
					'<w:r><w:rPr><w:b/></w:rPr><w:t>DOC-1</w:t></w:r>' +
					'<w:r><w:fldChar w:fldCharType="end"/></w:r>',
			),
		});
		const loaded = await loadDocx(original);
		const saved = await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-0042'));
		const header = (await parts(saved)).get('word/header1.xml')!;
		expect(header).toMatch(/<w:r><w:rPr><w:i\/><\/w:rPr><w:t[^>]*>DOC-0042<\/w:t><\/w:r>/);
		expect(header).not.toContain('<w:b/>');
	});

	it('saves without the refresh when the header cannot be rebuilt', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				// Bidirectional override text: inline markup the writer cannot rebuild.
				`${complexField('STYLEREF "Doc Number"', ['DOC-1'])}<w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>`,
			),
		});
		const loaded = await loadDocx(original);
		const saved = await parts(await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-2')));
		expect(saved.get('word/document.xml')).toContain('>DOC-2<');
		expect(saved.get('word/header1.xml')).toBe((await parts(original)).get('word/header1.xml'));
	});

	it('refreshes the other parts when one header cannot be rebuilt', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				`${complexField('STYLEREF "Doc Number"', ['DOC-1'])}<w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>`,
			),
			footer: paragraph('', undefined, simpleField('STYLEREF &quot;Doc Number&quot;', 'DOC-1')),
		});
		const loaded = await loadDocx(original);
		const saved = await parts(await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-2')));
		expect(saved.get('word/header1.xml')).toBe((await parts(original)).get('word/header1.xml'));
		expect(saved.get('word/footer1.xml')).toContain('>DOC-2<');
		expect(saved.get('word/footer1.xml')).not.toContain('DOC-1');
	});

	it('reports the error of a header edit the writer cannot rebuild', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				`<w:r><w:t>Plan</w:t></w:r><w:bdo w:val="rtl"><w:r><w:t>x</w:t></w:r></w:bdo>${complexField('STYLEREF "Doc Number"', ['DOC-1'])}`,
			),
		});
		const loaded = await loadDocx(original);
		const model = setText(loaded.model, 'DocNumber', 'DOC-2');
		const header = expectParagraph(must(model.sections?.[0]?.headers?.default).blocks[0]);
		header.runs[0] = { ...header.runs[0]!, text: 'Plan B' };
		await expect(loaded.save(model)).rejects.toThrow(/Cannot edit paragraph/);
	});

	it('saves a refreshed result in the text a revision keeps', async () => {
		const original = await styleRefFixture({
			body: paragraph('DOC-1', 'DocNumber'),
			header: paragraph(
				'',
				undefined,
				'<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
					'<w:r><w:instrText xml:space="preserve"> STYLEREF "Doc Number" </w:instrText></w:r>' +
					'<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
					'<w:del w:id="1" w:author="A"><w:r><w:delText>DOC-0</w:delText></w:r></w:del>' +
					'<w:ins w:id="2" w:author="A"><w:r><w:t>DOC-1</w:t></w:r></w:ins>' +
					'<w:r><w:fldChar w:fldCharType="end"/></w:r>',
			),
		});
		const loaded = await loadDocx(original);
		const header = (
			await parts(await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-2')))
		).get('word/header1.xml')!;
		expect(header).toMatch(/<w:del\b[^>]*><w:r><w:delText>DOC-0<\/w:delText><\/w:r><\/w:del>/);
		expect(header).toMatch(/<w:ins\b[^>]*><w:r><w:t>DOC-2<\/w:t><\/w:r><\/w:ins>/);
		const reloaded = await loadDocx(await loaded.save(setText(loaded.model, 'DocNumber', 'DOC-2')));
		const runs = expectParagraph(
			must(reloaded.model.sections?.[0]?.headers?.default).blocks[0],
		).runs;
		expect(runs.filter((run) => run.field).map((run) => [run.text, run.revision?.kind])).toEqual([
			['DOC-0', 'delete'],
			['DOC-2', 'insert'],
		]);
	});
});
