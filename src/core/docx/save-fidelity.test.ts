import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx, type DocumentModel } from './index';
import { at, expectParagraph, must } from './test-support/access';
import { signedTwips } from './units';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
// Word starts every part with this declaration and a CRLF, and writes no folder entries.
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
// Word interleaves the references: the default header and footer, then the first-page ones.
const SECTION =
	'<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>' +
	'<w:headerReference w:type="first" r:id="rId3"/><w:footerReference w:type="first" r:id="rId4"/>' +
	'<w:pgSz w:w="11906" w:h="16838"/>' +
	'<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
	'<w:cols w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>';
const STORIES = [
	['rId1', 'header', 'header1', 'hdr'],
	['rId2', 'footer', 'footer1', 'ftr'],
	['rId3', 'header', 'header2', 'hdr'],
	['rId4', 'footer', 'footer2', 'ftr'],
] as const;

const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

/** A package shaped like one Word writes, around `body`, with first-page headers and footers. */
async function wordPackage(body: string, folders = false): Promise<Uint8Array> {
	const zip = new JSZip();
	const part = (name: string, xml: string) =>
		zip.file(name, DECLARATION + xml, { createFolders: folders });
	part(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	part(
		'_rels/.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	part(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${SECTION}</w:body></w:document>`,
	);
	part(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${RELS}">${STORIES.map(
			([id, type, name]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${name}.xml"/>`,
		).join('')}</Relationships>`,
	);
	for (const [, , name, root] of STORIES)
		part(`word/${name}.xml`, `<w:${root} xmlns:w="${W}"><w:p>${text(name)}</w:p></w:${root}>`);
	return zip.generateAsync({ type: 'uint8array' });
}

async function saved(
	loaded: Awaited<ReturnType<typeof loadDocx>>,
	edit: (model: DocumentModel) => void,
) {
	const model = structuredClone(loaded.model);
	edit(model);
	return JSZip.loadAsync(await loaded.save(model));
}

const part = (zip: JSZip, name: string) => must(zip.file(name), name).async('string');
const sectionOf = async (zip: JSZip) =>
	(await part(zip, 'word/document.xml')).match(/<w:sectPr>[\s\S]*<\/w:sectPr>/)?.[0];
const folders = (zip: JSZip) =>
	Object.values(zip.files)
		.filter((file) => file.dir)
		.map((file) => file.name);
const editBody = (model: DocumentModel) => {
	at(expectParagraph(model.blocks[0]).runs, 0).text = 'Edited';
};

describe('saving an edited Word package', () => {
	it('adds no folder entries to a package that has none', async () => {
		const loaded = await loadDocx(await wordPackage(`<w:p>${text('Body')}</w:p>`));
		const zip = await saved(loaded, editBody);
		expect(await part(zip, 'word/document.xml')).toContain('>Edited<');
		expect(folders(zip)).toEqual([]);
		const fresh = await JSZip.loadAsync(await saveDocx(structuredClone(loaded.model)));
		expect(folders(fresh)).toEqual([]);
	});

	it('keeps the folder entries a package already has', async () => {
		const source = await JSZip.loadAsync(await wordPackage(`<w:p>${text('Body')}</w:p>`, true));
		const loaded = await loadDocx(await wordPackage(`<w:p>${text('Body')}</w:p>`, true));
		expect(folders(await saved(loaded, editBody)).sort()).toEqual(folders(source).sort());
	});
	it('keeps the order of the header and footer references when the body changes', async () => {
		const loaded = await loadDocx(await wordPackage(`<w:p>${text('Body')}</w:p>`));
		expect(await sectionOf(await saved(loaded, editBody))).toBe(SECTION);
	});
	it('leaves the section properties as they were when only a header changes', async () => {
		const loaded = await loadDocx(await wordPackage(`<w:p>${text('Body')}</w:p>`));
		const zip = await saved(loaded, (model) => {
			const header = must(at(model.sections, 0).headers?.default, 'default header');
			at(expectParagraph(header.blocks[0]).runs, 0).text = 'Edited header';
		});
		expect(await part(zip, 'word/header1.xml')).toContain('>Edited header<');
		expect(await sectionOf(zip)).toBe(SECTION);
	});

	it('writes only the section properties that changed', async () => {
		const loaded = await loadDocx(await wordPackage(`<w:p>${text('Body')}</w:p>`));
		const zip = await saved(loaded, (model) => {
			at(model.sections, 0).marginTopTwips = signedTwips(1000);
		});
		expect(await sectionOf(zip)).toBe(SECTION.replace('w:top="1417"', 'w:top="1000"'));
	});
});
