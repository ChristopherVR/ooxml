// Synthetic samples for the tab stop reports (#414, #385, #404), built from the descriptions in the
// issues: A4, 1418 twip margins, Arial 11 pt. Run from this folder: node generate-tab-fixtures.mjs
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';

const styles = (extra = '') =>
	`<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${extra}</w:styles>`;
const text = (value) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const tab = '<w:r><w:tab/></w:r>';
const paragraph = (properties, ...runs) =>
	`<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${runs.join('')}</w:p>`;

async function build(file, body, stylesXml = styles(), margin = 1418) {
	const zip = new JSZip();
	const part = (name, xml) =>
		zip.file(name, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
	part(
		'[Content_Types].xml',
		`<Types xmlns="${types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${relationships}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/></Relationships>`,
	);
	part('word/styles.xml', stylesXml);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="${margin}" w:bottom="1418" w:left="${margin}" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(file, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}

// #414: the same two stops in a style and directly in the paragraph.
const stops =
	'<w:tabs><w:tab w:val="left" w:pos="4536"/><w:tab w:val="right" w:pos="9070"/></w:tabs>';
const row = (properties) =>
	paragraph(properties, text('Name'), tab, text('Value'), tab, text('Page 12'));
const tabbed = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="TabbedBody"><w:name w:val="Tabbed Body"/><w:basedOn w:val="Normal"/><w:pPr>${stops}</w:pPr></w:style>`;
const derived = `<w:style w:type="paragraph" w:customStyle="1" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="TabbedBody"/><w:pPr><w:tabs><w:tab w:val="clear" w:pos="4536"/><w:tab w:val="left" w:pos="2000"/></w:tabs></w:pPr></w:style>`;
await build(
	'issue-414-style-tab-stops.docx',
	row('<w:pStyle w:val="TabbedBody"/>') + row('<w:pStyle w:val="Derived"/>'),
	styles(tabbed + derived),
);
await build('issue-414-direct-tab-stops.docx', row(stops) + row(stops));

// #385: a right stop past the right margin (9354 twips against a 9070 twip text area).
const rightTab = (pos) => `<w:tabs><w:tab w:val="right" w:pos="${pos}"/></w:tabs>`;
const rows = (pos) =>
	[
		paragraph(rightTab(pos) + '<w:jc w:val="both"/>', tab, text('Total 1,250.00 EUR')),
		paragraph(rightTab(pos), text('Amount due'), tab, text('1,250.00 EUR')),
		paragraph(rightTab(pos), text('Invoice 2026-114'), tab, text('Page 12 of 30')),
	].join('');
await build('issue-385-right-tab-past-margin.docx', rows(9354));
await build('issue-385-right-tab-at-margin.docx', rows(9070));

// #404: text after a right tab at the margin (9072 twips, margins 1417), and a right ptab.
const atMargin = (pos) => `<w:tabs><w:tab w:val="right" w:pos="${pos}"/></w:tabs>`;
const ptab = '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>';
const labelled = (properties, middle, value) =>
	paragraph(properties, text('Label'), middle, text(value));
const values = ['Page 1', 'Some words here', 'Total 12.50', 'Hello world', 'Item 3 of 9'];
await build(
	'issue-404-tab-stop-at-margin.docx',
	values.map((value) => labelled(atMargin(9072), tab, value)).join('') +
		values.map((value) => labelled('', ptab, value)).join(''),
	styles(),
	1417,
);
await build(
	'issue-404-tab-stop-one-twip-inside.docx',
	values.map((value) => labelled(atMargin(9071), tab, value)).join(''),
	styles(),
	1417,
);

// #396: 10 pt Arial lines with a dot leader to 8931 twips; the dots of all lines share one grid.
const leaderStops =
	'<w:tabs><w:tab w:val="left" w:pos="2835"/><w:tab w:val="left" w:leader="dot" w:pos="8931"/></w:tabs>';
const onlyLeader = '<w:tabs><w:tab w:val="left" w:leader="dot" w:pos="8931"/></w:tabs>';
const small = (value) => `<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>${value}</w:t></w:r>`;
const tabSmall = '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:tab/></w:r>';
await build(
	'issue-396-tab-leader-dots.docx',
	paragraph(leaderStops, small('Name'), tabSmall, tabSmall) +
		paragraph(leaderStops, small('Date and place of birth'), tabSmall, tabSmall) +
		paragraph(onlyLeader, small('Name'), tabSmall) +
		paragraph(onlyLeader, small('Postal address'), tabSmall),
);

// #420: text after a left tab whose first word does not fit after the last stop (margins 1417).
const longWords = 'Extraordinarily long words follow';
const twoStops = (first) =>
	`<w:tabs><w:tab w:val="left" w:pos="${first}"/><w:tab w:val="left" w:pos="8000"/></w:tabs><w:spacing w:after="240"/>`;
const stopsRows = (value) =>
	[1134, 3000]
		.map((first) =>
			paragraph(twoStops(first), text(`Stops at ${first} and 8000 twips`), tab, text(value)),
		)
		.join('');
await build('issue-420-left-tab-text-does-not-fit.docx', stopsRows(longWords), styles(), 1417);
await build('issue-420-left-tab-text-fits-control.docx', stopsRows('Short'), styles(), 1417);
