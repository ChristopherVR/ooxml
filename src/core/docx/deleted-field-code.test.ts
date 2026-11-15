import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { acceptAllRevisions, rejectAllRevisions } from './revision-commands';
import { expectParagraph } from './test-support/access';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const field = (code: string) =>
	`<w:r><w:fldChar w:fldCharType="begin"/></w:r>${code}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

async function load(paragraph: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${paragraph}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
async function documentXml(bytes: Uint8Array): Promise<string> {
	return (await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) ?? '';
}

describe('deleted field codes (w:delInstrText)', () => {
	it('reads a deleted field code and writes it back as w:delInstrText', async () => {
		const loaded = await load(
			`<w:p>${field('<w:del w:id="1" w:author="A"><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r></w:del><w:ins w:id="2" w:author="A"><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r></w:ins>')}</w:p>`,
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(
			paragraph.runs
				.filter((run) => run.fieldCode !== undefined)
				.map((run) => [run.fieldCode, run.revision?.kind]),
		).toEqual([
			[' PAGE ', 'delete'],
			[' NUMPAGES ', 'insert'],
		]);
		const rejected = await documentXml(await loaded.save(rejectAllRevisions(loaded.model)));
		expect(rejected).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
		expect(rejected).not.toContain('NUMPAGES');
		expect(rejected).not.toContain('delInstrText');
		const accepted = await documentXml(await loaded.save(acceptAllRevisions(loaded.model)));
		expect(accepted).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>');
		expect(accepted).not.toContain(' PAGE ');
	});

	it('writes w:delInstrText for a field code the model marks deleted', async () => {
		const loaded = await load(
			`<w:p>${field('<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>')}</w:p>`,
		);
		const model = structuredClone(loaded.model);
		const run = expectParagraph(model.blocks[0]).runs.find((item) => item.fieldCode !== undefined)!;
		run.revision = { kind: 'delete', author: 'A', id: 'new' };
		const xml = await documentXml(await loaded.save(model));
		expect(xml).toMatch(
			/<w:del [^>]*><w:r><w:delInstrText xml:space="preserve"> PAGE <\/w:delInstrText><\/w:r><\/w:del>/,
		);
	});
});
