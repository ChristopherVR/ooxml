// Paragraphs as Word saves them after proofing: spelling marks, a permission range, a bookmark,
// a smart tag, a symbol character and a positional tab between ordinary runs.
import JSZip from 'jszip';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

export async function inlineMarkersFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>` +
			`<w:p>${run('A ')}<w:proofErr w:type="spellStart"/>${run('speling')}<w:proofErr w:type="spellEnd"/>` +
			`<w:permStart w:id="1" w:edGrp="everyone"/>${run(' mistake')}<w:permEnd w:id="1"/></w:p>` +
			`<w:p><w:bookmarkStart w:id="2" w:name="Place"/>${run('Meet in ')}` +
			`<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="City">${run('Paris')}</w:smartTag>` +
			`<w:bookmarkEnd w:id="2"/></w:p>` +
			`<w:p>${run('Approved')}<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>` +
			`<w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r></w:p>` +
			`<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
