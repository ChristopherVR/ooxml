// MERGEFIELD in headers and footers of the Word PDF export (issue #368).
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { exportDocxPdf } from './export-pdf';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The standard-font text a PDF shows, in drawing order. */
function pdfText(bytes: Uint8Array): string {
	const raw = Buffer.from(bytes).toString('latin1');
	let text = '';
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (!/\/Filter \/FlateDecode/.test(match[1]!) || /\/Subtype \/Image|\/Length1/.test(match[1]!))
			continue;
		const content = inflateSync(bytes.subarray(start, start + length)).toString('latin1');
		for (const shown of content.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g))
			text += shown[1]!.replace(/\\(.)/g, '$1');
	}
	return text;
}

// Word's markup: a complex MERGEFIELD then a simple one, each with a cached result.
const fields = (label: string) =>
	`<w:p><w:r><w:t xml:space="preserve">${label}: </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD  Reference  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Cached A</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> - </w:t></w:r><w:fldSimple w:instr=" MERGEFIELD Name \\* MERGEFORMAT"><w:r><w:t>Cached B</w:t></w:r></w:fldSimple></w:p>`;

async function headerDocument(settings?: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${fields('Body')}<w:sectPr><w:headerReference w:type="default" r:id="hdr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}">${fields('Header')}</w:hdr>`);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="hdr1" Type="${REL}/header" Target="header1.xml"/>${settings ? `<Relationship Id="set" Type="${REL}/settings" Target="settings.xml"/>` : ''}</Relationships>`,
	);
	if (settings)
		zip.file('word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

describe('exportDocxPdf merge fields in headers and footers', () => {
	it('shows «Name» in the header of a document that is not a mail merge', async () => {
		// Word 16: "Header: «Reference» - «Name»" over "Body: Cached A - Cached B".
		const model = await headerDocument();
		expect(model.mailMerge).toBeUndefined();
		const text = pdfText((await exportDocxPdf(model, { measurer: createFakeMeasurer() })).bytes);
		expect(text).toContain('Header: \xabReference\xbb - \xabName\xbb');
		expect(text).toContain('Body: Cached A - Cached B');
		expect(text).not.toContain('Header: Cached');
	});

	it('keeps the cached results in the header of a mail merge main document', async () => {
		const model = await headerDocument(
			'<w:mailMerge><w:mainDocumentType w:val="formLetters"/><w:dataType w:val="textFile"/><w:query w:val="SELECT * FROM x"/></w:mailMerge>',
		);
		expect(model.mailMerge).toBeDefined();
		const text = pdfText((await exportDocxPdf(model, { measurer: createFakeMeasurer() })).bytes);
		expect(text).toContain('Header: Cached A - Cached B');
	});
});
