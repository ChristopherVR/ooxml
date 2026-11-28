// Synthetic reproductions of #423 (footnote that does not fit at the page end) and #424 (row
// minimum height at the page end). The reporter's files are not available, so these follow the
// parameters stated in the issues. Run with `node`.
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';

async function docx(name, body, { font, size, footnotes = '' }) {
	const zip = new JSZip();
	const part = (path, xml) =>
		zip.file(path, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	const rel = (id, type, target) =>
		`<Relationship Id="${id}" Type="${office}/${type}" Target="${target}"/>`;
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${footnotes ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : ''}</Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}">${rel('rId1', 'styles', 'styles.xml')}${footnotes ? rel('rId2', 'footnotes', 'footnotes.xml') : ''}</Relationships>`,
	);
	part(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/><w:sz w:val="${size}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>`,
	);
	if (footnotes) part('word/footnotes.xml', footnotes);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}
const p = (text, pPr = '', rPr = '') =>
	`<w:p>${pPr && `<w:pPr>${pPr}</w:pPr>`}<w:r>${rPr && `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const filler = (count) =>
	Array.from({ length: count }, (_, i) => p(`Filler line ${i + 1}`)).join('');

// #424: 48 lines of Times New Roman 12 leave about 37.8 pt above the bottom margin, then a
// one-row table of five 10 pt lines, with or without a minimum height.
const table = (height) =>
	`<w:tbl><w:tblPr><w:tblW w:w="3614" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3614"/></w:tblGrid><w:tr>${height ? `<w:trPr><w:trHeight w:val="${height}"/></w:trPr>` : ''}<w:tc><w:tcPr><w:tcW w:w="3614" w:type="dxa"/></w:tcPr>${[1, 2, 3, 4, 5].map((n) => p(`Cell line ${n}`, '', '<w:sz w:val="20"/>')).join('')}</w:tc></w:tr></w:tbl>${p('After the table')}`;
const tnr = { font: 'Times New Roman', size: 24 };
await docx('issue-424-row-min-height-at-page-end', filler(48) + table(900), tnr);
await docx('issue-424-row-no-min-height-control', filler(48) + table(0), tnr);

// #423: 44 lines of Calibri 11, then a reference line whose 15-line note does not fit.
const note = (lines) =>
	`<w:footnotes xmlns:w="${w}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${Array.from({ length: lines * 12 }, (_, i) => `note${i + 1}`).join(' ')}</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const withNote = `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t xml:space="preserve">Paragraph with notes: articles 54</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t xml:space="preserve"> of the code.</w:t></w:r></w:p>`;
const calibri = { font: 'Calibri', size: 22 };
await docx('issue-423-long-footnote-split', filler(44) + withNote, {
	...calibri,
	footnotes: note(15),
});
await docx('issue-423-short-footnote-control', filler(44) + withNote, {
	...calibri,
	footnotes: note(3),
});

// #424, text that fits but whose minimum height does not: bordered rows, 47 filler lines.
const bordered = (height) =>
	`<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => `<w:${n} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${height ? `<w:trPr><w:trHeight w:val="${height}"/></w:trPr>` : ''}<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${[1, 2].map((n) => p(`Text line ${n}`, '<w:spacing w:before="100" w:after="0"/>', '<w:sz w:val="20"/>')).join('')}</w:tc></w:tr></w:tbl>${p('After the table')}`;
await docx('issue-424-row-min-height-content-fits', filler(47) + bordered(1453), tnr);
await docx('issue-424-row-content-fits-control', filler(47) + bordered(0), tnr);
