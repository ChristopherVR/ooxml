// Regenerates the synthetic line-breaking fixtures of #406, #410 and #392 from the
// paragraphs quoted in those issues (the reporter's Word files are not redistributed here).
// Run: node generate-line-break-fixtures.mjs
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const w15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';

async function build(name, body, { font, size, bold = false }) {
	const zip = new JSZip();
	const part = (path, xml) =>
		zip.file(path, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${office}/settings" Target="settings.xml"/></Relationships>`,
	);
	part(
		'word/settings.xml',
		`<w:settings xmlns:w="${w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
	);
	part(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${bold ? '<w:b/>' : ''}<w:sz w:val="${size * 2}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
	);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:w15="${w15}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

const para = (runs, jc = '') =>
	`<w:p><w:pPr><w:spacing w:after="120"/>${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${runs}</w:p>`;
const run = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

// #406: a space next to a non-breaking space is no break opportunity.
const starts = [
	'The',
	'The report',
	'The first report',
	'The first full report',
	'The first and full report',
	'The first and only full report',
];
const tail = 'of the committee and the expert opinion were both written by Dr John';
const rest = 'SMITH, specialist.';
const nbspBody = (between) => starts.map((start) => para(run(`${start} ${tail}${between}${rest}`)));
await build(
	'issue-406-space-next-to-nbsp',
	[...nbspBody('&#xA0; '), ...nbspBody(' &#xA0;')].join(''),
	{ font: 'Times New Roman', size: 12 },
);
await build('issue-406-two-spaces-control', nbspBody('  ').join(''), {
	font: 'Times New Roman',
	size: 12,
});

// #410: content controls shown as tags.
const tagged = (alias, id, inner) =>
	`<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${alias}"/><w:id w:val="${id}"/><w15:appearance w15:val="tags"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const nested = (base) =>
	tagged(
		'Outer',
		base + 2,
		`${tagged('Inner', base + 1, '<w:r><w:t>Value</w:t></w:r>')}<w:r><w:t xml:space="preserve"> text</w:t></w:r>`,
	);
const after =
	' alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau.';
const tagParagraphs = (head, field) =>
	['both', 'left']
		.map(
			(jc) =>
				`<w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>${head}${field}${nested(4000)}${run(after)}</w:p>`,
		)
		.join('');
const word410 = (space) => run(`${'i'.repeat(67)} lorem ipsum dolor sit amet Name/Field${space}`);
await build('issue-410-start-tag-after-word', tagParagraphs(word410(''), ''), {
	font: 'Arial',
	size: 11,
});
await build('issue-410-start-tag-after-space', tagParagraphs(word410(' '), ''), {
	font: 'Arial',
	size: 11,
});
// #392: Word draws spaces in small caps at the reduced size.
const smallCaps = (on) =>
	para(
		`<w:r><w:rPr>${on ? '<w:smallCaps/>' : ''}</w:rPr><w:t xml:space="preserve">${Array(26).fill('AB').join(' ')} NEXT LINE</w:t></w:r>`,
	);
await build('issue-392-small-caps-spaces', smallCaps(true), {
	font: 'Times New Roman',
	size: 11,
	bold: true,
});
await build('issue-392-small-caps-spaces-control', smallCaps(false), {
	font: 'Times New Roman',
	size: 11,
	bold: true,
});
