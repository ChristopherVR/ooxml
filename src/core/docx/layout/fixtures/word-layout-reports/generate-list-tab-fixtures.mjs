// Synthetic samples for #421 (the number takes the paragraph mark's formatting) and #422 (the tab
// after a number stops at the hanging indent), built from the issues' descriptions: A4, margins
// 1417 twips, Arial 11 pt. Run from this folder: node generate-list-tab-fixtures.mjs
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';
const styles = `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
const level = (format, text, indent, rPr = '') =>
	`<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr>${indent}</w:pPr>${rPr}</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

async function build(file, body, numbering) {
	const zip = new JSZip();
	const part = (name, xml) =>
		zip.file(name, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${office}/numbering" Target="numbering.xml"/></Relationships>`,
	);
	part('word/styles.xml', styles);
	part('word/numbering.xml', numbering);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1417" w:bottom="1418" w:left="1417" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(file, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

// #421: decimal list without level formatting; the mark and the first run differ.
const num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
const decimal = level('decimal', '%1.', '<w:ind w:left="720" w:hanging="360"/>');
const run = (rPr, value) => `<w:r><w:rPr>${rPr}</w:rPr><w:t>${value}</w:t></w:r>`;
const item = (markRPr, runRPr, value) =>
	`<w:p><w:pPr>${num}${markRPr ? `<w:rPr>${markRPr}</w:rPr>` : ''}</w:pPr>${run(runRPr, value)}</w:p>`;
await build(
	'issue-421-number-from-first-run.docx',
	item('<w:sz w:val="16"/>', '<w:sz w:val="40"/>', 'Mark 8 pt, text 20 pt') +
		item('', '<w:sz w:val="40"/>', 'No mark formatting, text 20 pt') +
		item(
			'<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:b/>',
			'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:i/>',
			'Mark Courier bold, text Arial italic',
		) +
		item('<w:sz w:val="40"/>', '<w:sz w:val="16"/>', 'Mark 20 pt, text 8 pt') +
		item('', '', 'Next paragraph'),
	decimal,
);
await build(
	'issue-421-number-mark-control.docx',
	item('<w:sz w:val="16"/>', '<w:sz w:val="16"/>', 'Mark 8 pt, text 8 pt') +
		item('<w:sz w:val="40"/>', '<w:sz w:val="40"/>', 'Mark 20 pt, text 20 pt') +
		item('', '', 'Next paragraph'),
	decimal,
);

// #422: a bullet whose paragraph has one custom left stop at a given position.
const bullet = level(
	'bullet',
	'-',
	'<w:ind w:left="384" w:hanging="142"/>',
	'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>',
);
const bulleted = (stop) =>
	`<w:p><w:pPr>${num}${stop ? `<w:tabs><w:tab w:val="left" w:pos="${stop}"/></w:tabs>` : ''}</w:pPr><w:r><w:t>Text after the bullet</w:t></w:r></w:p>`;
await build(
	'issue-422-list-tab-past-indent.docx',
	[200, 340, 500, 1235].map(bulleted).join(''),
	bullet,
);
await build('issue-422-list-no-custom-tab-control.docx', bulleted(0), bullet);
