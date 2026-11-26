import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
const types = 'http://schemas.openxmlformats.org/package/2006/content-types';
const zip = new JSZip();
const part = (name, xml) =>
	zip.file(name, xml, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
const p = (properties = '') =>
	`<w:p>${properties && `<w:pPr>${properties}</w:pPr>`}<w:r><w:t>Lorem ipsum.</w:t></w:r></w:p>`;

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
	`<w:styles xmlns:w="${w}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:jc w:val="left"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Custom"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style><w:style w:type="table" w:styleId="Example"><w:name w:val="Example"/><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:tblStylePr w:type="firstRow"><w:pPr><w:spacing w:before="80"/></w:pPr></w:tblStylePr></w:style></w:styles>`,
);
part(
	'word/document.xml',
	`<w:document xmlns:w="${w}"><w:body>${p()}<w:tbl><w:tblPr><w:tblStyle w:val="Example"/><w:tblLook w:firstRow="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid><w:tr><w:tc>${p()}${p('<w:pStyle w:val="Normal"/>')}${p('<w:pStyle w:val="Custom"/>')}${p('<w:spacing w:after="360"/><w:jc w:val="right"/>')}</w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
);
await writeFile(
	new URL('./issue-353-table-style-precedence.docx', import.meta.url),
	await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
);
