import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { must } from './test-support/access';
import { signedTwips, twips } from './units';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const STYLES = `<w:styles xmlns:w="${ns}" xmlns:w15="${W15}"><w:style w:type="paragraph" w:styleId="Body" w:default="1"><w:name w:val="Normal"/><w:uiPriority w:val="1"/><w:pPr><w:spacing w:after="240"/><w:textAlignment w:val="baseline"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:qFormat/></w:style></w:styles>`;
const DOC = `<w:document xmlns:w="${ns}"><w:body><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

async function load() {
	const zip = new JSZip();
	zip.file('word/document.xml', DOC);
	zip.file('word/styles.xml', STYLES);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const stylesOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/styles.xml')!.async('string')) as string;

describe('editing the formatting of existing paragraph styles', () => {
	it('writes changed properties and keeps unknown pPr, rPr and style children', async () => {
		const loaded = await load();
		const body = must(loaded.model.paragraphStyles?.styles.Body, 'Body style');
		body.formatting.spacingAfterTwips = twips(120);
		body.formatting.align = 'center';
		body.formatting.keepNext = true;
		body.formatting.indentLeftTwips = signedTwips(360);
		const xml = await stylesOf(await loaded.save(loaded.model));
		expect(xml).toContain('<w:spacing w:after="120"/>');
		expect(xml).toContain('<w:jc w:val="center"/>');
		expect(xml).toContain('<w:keepNext/>');
		expect(xml).toContain('<w:ind w:left="360"/>');
		expect(xml).toContain('<w:textAlignment w:val="baseline"/>');
		expect(xml).toContain('<w:uiPriority w:val="1"/><w:pPr>');
		expect(xml).toContain('<w:rPr><w:b/></w:rPr>');
		const reopened = must((await loadDocx(await loaded.save(loaded.model))).model.paragraphStyles);
		expect(reopened.styles.Body?.formatting).toMatchObject({
			spacingAfterTwips: 120,
			align: 'center',
			keepNext: true,
			indentLeftTwips: 360,
		});
	});

	it('adds a pPr in schema order to a style without one and removes a property', async () => {
		const loaded = await load();
		must(loaded.model.paragraphStyles?.styles.Plain).formatting.pageBreakBefore = true;
		delete must(loaded.model.paragraphStyles?.styles.Body).formatting.spacingAfterTwips;
		const xml = await stylesOf(await loaded.save(loaded.model));
		expect(xml).toContain('<w:qFormat/><w:pPr><w:pageBreakBefore/></w:pPr></w:style>');
		expect(xml).not.toContain('w:after="240"');
		expect(xml).toContain('<w:textAlignment w:val="baseline"/>');
	});

	it('still refuses renames and outline level edits', async () => {
		const renamed = await load();
		must(renamed.model.paragraphStyles?.styles.Body).name = 'Other';
		await expect(renamed.save(renamed.model)).rejects.toThrow(/style catalog is not supported/);
		const outlined = await load();
		must(outlined.model.paragraphStyles?.styles.Body).formatting.outlineLevel = 2;
		await expect(outlined.save(outlined.model)).rejects.toThrow(/style catalog is not supported/);
	});
});
