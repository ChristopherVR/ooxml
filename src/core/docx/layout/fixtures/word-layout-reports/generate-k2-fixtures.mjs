// Regenerates the synthetic fixtures of #429, #431 and #433 from the XML quoted in those issues
// (the reporter's Word files are not redistributed here). Run: node generate-k2-fixtures.mjs
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';
const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

async function build(name, body, { mode, numbering = '' }) {
	const zip = new JSZip();
	const part = (path, xml) =>
		zip.file(path, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	const override = (path, type) =>
		`<Override PartName="${path}" ContentType="${wml}.${type}+xml"/>`;
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${override('/word/document.xml', 'document.main')}${override('/word/styles.xml', 'styles')}${override('/word/settings.xml', 'settings')}${numbering ? override('/word/numbering.xml', 'numbering') : ''}</Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${office}/settings" Target="settings.xml"/>${numbering ? `<Relationship Id="rId3" Type="${office}/numbering" Target="numbering.xml"/>` : ''}</Relationships>`,
	);
	part(
		'word/settings.xml',
		`<w:settings xmlns:w="${w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat></w:settings>`,
	);
	part(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
	);
	if (numbering) part('word/numbering.xml', numbering);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

const text = (t) => `<w:r><w:t>${t}</w:t></w:r>`;
const para = (inner, ppr = '') => `<w:p><w:pPr>${ppr}</w:pPr>${inner}</w:p>`;

// #429: a text-wrapping break that ends a paragraph, in its own run and after text in one run.
const lines = [
	'Lorem ipsum dolor sit amet.',
	'Consectetur adipiscing elit.',
	'Sed do eiusmod tempor.',
	'Ut enim ad minim veniam.',
	'Quis nostrud exercitation.',
];
const br = '<w:br/>';
await build(
	'issue-429-line-break-at-paragraph-end',
	para(text(lines[0]) + `<w:r>${br}</w:r>`) +
		para(text(lines[1])) +
		para(`<w:r><w:t>${lines[2]}</w:t>${br}</w:r>`) +
		para(text(lines[3])) +
		para(text(lines[4])),
	{ mode: 15 },
);
await build(
	'issue-429-line-break-at-paragraph-end-control',
	lines.map((l) => para(text(l))).join(''),
	{
		mode: 15,
	},
);

// #431: a paragraph with space before that starts with a page break in the same run as its text.
const before = '<w:spacing w:before="388" w:after="0"/>';
const pageBreakBody =
	para(text('Lorem ipsum dolor sit amet.')) +
	para(`<w:r><w:br w:type="page"/><w:t>Consectetur adipiscing elit.</w:t></w:r>`, before) +
	para(text('Sed do eiusmod tempor.'));
await build('issue-431-page-break-space-before-compat12', pageBreakBody, { mode: 12 });
await build('issue-431-page-break-space-before-compat15-control', pageBreakBody, { mode: 15 });

// #433: a bullet whose level text is a soft hyphen, and the same with a hyphen-minus.
const bullets = (lvlText) =>
	`<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${lvlText}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
const items = ['Alpha', 'Beta', '', 'Gamma']
	.map((t) => para(t ? text(t) : '', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
	.join('');
await build('issue-433-bullet-soft-hyphen', items, { mode: 15, numbering: bullets('&#xAD;') });
await build('issue-433-bullet-hyphen-minus-control', items, { mode: 15, numbering: bullets('-') });
