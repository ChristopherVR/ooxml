import JSZip from 'jszip';
export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function fieldDocx(simple = false, adjacent = false): Promise<Buffer> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t xml:space="preserve">Author: </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> AUTHOR </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Ann</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> wrote this.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
	);
	if (simple) {
		const xml = await zip.file('word/document.xml')!.async('string');
		zip.file(
			'word/document.xml',
			xml.replace(
				/<w:r><w:fldChar w:fldCharType="begin"\/><\/w:r>[\s\S]*?<w:r><w:fldChar w:fldCharType="end"\/><\/w:r>/,
				'<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann</w:t></w:r></w:fldSimple>'.repeat(
					adjacent ? 2 : 1,
				),
			),
		);
	}
	return zip.generateAsync({ type: 'nodebuffer' });
}
