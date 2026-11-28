// Synthetic reproductions of the table geometry reports #387, #405, #407 and #413 (the reporter's
// files are not available, so these follow the parameters stated in the issues). Run with `node`.
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';

async function docx(name, body) {
	const zip = new JSZip();
	const part = (path, xml) =>
		zip.file(path, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
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
	part(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style></w:styles>`,
	);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	await writeFile(
		new URL(`./${name}.docx`, import.meta.url),
		await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
	);
}
const p = (text, pPr = '') =>
	`<w:p>${pPr && `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (width, text, pPr = '', tcPr = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${tcPr}</w:tcPr>${p(text, pPr)}</w:tc>`;
const bd = (names, size) =>
	names.map((n) => `<w:${n} w:val="single" w:sz="${size}" w:space="0" w:color="000000"/>`).join('');

// #387: tables wider than the text area (9072 twips) keep their alignment.
const wide = (jc, width) =>
	`<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:jc w:val="${jc}"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="${(width / 2 - 0.5) | 0}"/><w:gridCol w:w="${width - ((width / 2 - 0.5) | 0)}"/></w:tblGrid><w:tr>${cell((width / 2 - 0.5) | 0, 'First')}${cell(width - ((width / 2 - 0.5) | 0), 'Second')}</w:tr></w:tbl>${p('after')}`;
await docx(
	'wide-tables-centered-right',
	wide('center', 9495) + wide('center', 10500) + wide('right', 9495),
);
await docx('narrow-tables-control', wide('center', 9000) + wide('right', 8000));

// #405: trHeight atLeast 1000 (50 pt) in a bordered table.
const rows = (size, borders) => {
	const tbl = borders
		? `<w:tblBorders>${bd(['top', 'left', 'bottom', 'right', 'insideH', 'insideV'], size)}</w:tblBorders>`
		: `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((n) => `<w:${n} w:val="none"/>`).join('')}</w:tblBorders>`;
	return `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${tbl}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${[1, 2, 3, 4].map((n) => `<w:tr><w:trPr><w:trHeight w:val="1000"/></w:trPr>${cell(2000, 'Row ' + n)}${cell(2000, 'x')}</w:tr>`).join('')}</w:tbl>${p('after')}`;
};
await docx('row-height-at-least-with-borders-sz6', rows(6, true));
await docx('row-height-at-least-with-borders-sz24', rows(24, true));
await docx('row-height-at-least-no-borders', rows(6, false));

// #413: horizontal borders, with or without insideV.
const grid = (tblBorders, tcBorders = '') =>
	`<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/>${tblBorders && `<w:tblBorders>${tblBorders}</w:tblBorders>`}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${[1, 2].map((n) => `<w:tr>${cell(2000, 'R' + n + 'A', '', tcBorders)}${cell(2000, 'R' + n + 'B', '', tcBorders)}</w:tr>`).join('')}</w:tbl>${p('after')}`;
const outer = ['top', 'left', 'bottom', 'right', 'insideH'];
await docx('table-borders-no-insidev', grid(bd(outer, 24)));
await docx('table-borders-with-insidev', grid(bd([...outer, 'insideV'], 24)));
await docx('table-borders-no-insidev-thin', grid(bd(outer, 4)));
await docx(
	'cell-borders-top-bottom',
	grid('', `<w:tcBorders>${bd(['top', 'bottom'], 24)}</w:tcBorders>`),
);

// #407: the fourth row starts page two; its paragraphs have spacing before.
const tall = (before) =>
	`<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>${bd([...outer, 'insideV'], 4)}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${[1, 2, 3, 4, 5].map((n) => `<w:tr><w:trPr><w:trHeight w:val="4600"/></w:trPr>${cell(4000, 'Row ' + n, `<w:spacing w:before="${before}" w:after="0"/>`)}</w:tr>`).join('')}</w:tbl>${p('after')}`;
await docx('row-at-page-top-space-before', tall(100));
await docx('row-at-page-top-no-space-before', tall(0));
