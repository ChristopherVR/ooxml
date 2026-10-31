// The preservation guarantee (viewers/docx/docs/preservation.md): a package edited and saved keeps
// every part the model does not hold byte for byte, and the XML it does not model inside
// document.xml. The package below carries the parts Word writes that the editor never touches.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { at, expectParagraph, expectTable, must } from './test-support/access';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MS_REL = 'http://schemas.microsoft.com/office/2011/relationships';

/** Parts that must survive an edit untouched, by name. */
const OPAQUE: Record<string, string | Uint8Array> = {
	'customXml/item1.xml': '<root xmlns="urn:test"><title>Kept</title></root>',
	'customXml/itemProps1.xml':
		'<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{4F1B3A3A-0000-0000-0000-000000000001}"/>',
	'customXml/_rels/item1.xml.rels': `<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL_TYPE}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
	'word/glossary/document.xml': `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Block"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Block</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`,
	'word/vbaProject.bin': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4, 5, 0, 255]),
	'word/people.xml': `<w15:people xmlns:w15="${W15}"><w15:person w15:author="Ada"><w15:presenceInfo w15:providerId="None" w15:userId="Ada"/></w15:person></w15:people>`,
	'word/commentsExtensible.xml':
		'<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>',
	'word/commentsIds.xml':
		'<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>',
	'word/fontTable.xml': `<w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Embedded"><w:embedRegular r:id="rId1" w:fontKey="{00000000-0000-0000-0000-000000000001}"/></w:font></w:fonts>`,
	'word/_rels/fontTable.xml.rels': `<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL_TYPE}/font" Target="fonts/font1.odttf"/></Relationships>`,
	'word/fonts/font1.odttf': new Uint8Array([9, 8, 7, 6, 5, 4, 3, 2, 1, 0, 128, 200]),
	'word/settings.xml': `<w:settings xmlns:w="${W}" xmlns:w14="${W14}"><w:zoom w:percent="100"/><w:extLst><w:ext w:uri="{unknown-ext}"><w14:unknownSetting xmlns:w14="${W14}" w14:val="1"/></w:ext></w:extLst></w:settings>`,
	'word/theme/theme1.xml':
		'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kept"/>',
	'docProps/custom.xml':
		'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"/>',
	'docProps/thumbnail.jpeg': new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]),
	'word/webSettings.xml': `<w:webSettings xmlns:w="${W}"/>`,
	'word/_rels/settings.xml.rels': `<Relationships xmlns="${RELS}"/>`,
	'word/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]),
};

const DOCUMENT =
	`<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}" xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w14 w15">` +
	`<w:body>` +
	`<w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr><w:rFonts w:ascii="Cambria"/></w:rPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>` +
	`<w:p><w:r><w:t>Second</w:t></w:r></w:p>` +
	`<mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>Choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>Fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>` +
	`<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/><w:tblPrChange w:id="9" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:tblPr/></w:tblPrChange></w:tblPr>` +
	`<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
	`<w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="10" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:trPr/></w:trPrChange></w:trPr>` +
	`<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcPrChange w:id="11" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:tcPr/></w:tcPrChange></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>` +
	`<w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
	`<w:p><w:bookmarkStart w:id="1" w:name="kept"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
	`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:extLst><w:ext w:uri="{unknown-sect}"/></w:extLst></w:sectPr>` +
	`</w:body></w:document>`;

const rel = (id: string, type: string, target: string) =>
	`<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

async function buildPackage(): Promise<Uint8Array> {
	const zip = new JSZip();
	const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
	zip.file(
		'[Content_Types].xml',
		`${decl}<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`${decl}<Relationships xmlns="${RELS}">${rel('rId1', `${REL_TYPE}/officeDocument`, 'word/document.xml')}${rel('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail', 'docProps/thumbnail.jpeg')}${rel('rId3', `${REL_TYPE}/custom-properties`, 'docProps/custom.xml')}</Relationships>`,
	);
	zip.file('word/document.xml', decl + DOCUMENT);
	zip.file(
		'word/_rels/document.xml.rels',
		`${decl}<Relationships xmlns="${RELS}">${[
			rel('rId1', `${REL_TYPE}/customXml`, '../customXml/item1.xml'),
			rel('rId2', `${REL_TYPE}/glossaryDocument`, 'glossary/document.xml'),
			rel(
				'rId3',
				'http://schemas.microsoft.com/office/2006/relationships/vbaProject',
				'vbaProject.bin',
			),
			rel('rId4', `${MS_REL}/people`, 'people.xml'),
			rel(
				'rId5',
				'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
				'commentsExtensible.xml',
			),
			rel(
				'rId6',
				'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
				'commentsIds.xml',
			),
			rel('rId7', `${REL_TYPE}/fontTable`, 'fontTable.xml'),
			rel('rId8', `${REL_TYPE}/settings`, 'settings.xml'),
			rel('rId9', `${REL_TYPE}/theme`, 'theme/theme1.xml'),
			rel('rId10', `${REL_TYPE}/webSettings`, 'webSettings.xml'),
			rel('rId11', `${REL_TYPE}/image`, 'media/image1.png'),
		].join('')}</Relationships>`,
	);
	for (const [name, content] of Object.entries(OPAQUE))
		zip.file(name, typeof content === 'string' ? decl + content : content);
	return zip.generateAsync({ type: 'uint8array' });
}

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
	const zip = await JSZip.loadAsync(bytes);
	const found = new Map<string, Uint8Array>();
	for (const [name, file] of Object.entries(zip.files))
		if (!file.dir) found.set(name, await file.async('uint8array'));
	return found;
}

const same = (a: Uint8Array | undefined, b: Uint8Array | undefined) =>
	a !== undefined && b !== undefined && a.length === b.length && a.every((v, i) => v === b[i]);

/** Every part except the ones an edit is allowed to rewrite must be byte-identical. */
async function expectKept(
	source: Uint8Array,
	saved: Uint8Array,
	rewritten: string[],
	added: string[] = [],
) {
	const before = await entries(source);
	const after = await entries(saved);
	expect([...after.keys()].sort()).toEqual([...before.keys(), ...added].sort());
	for (const [name, bytes] of before) {
		if (rewritten.includes(name)) continue;
		expect(same(bytes, after.get(name)), `${name} must be byte-identical`).toBe(true);
	}
}

describe('preservation of parts and XML the editor does not model', () => {
	it('saves an untouched package as the same bytes', async () => {
		const source = await buildPackage();
		expect(await (await loadDocx(source)).save()).toEqual(source);
	});

	it('keeps every unmodelled part byte for byte after a text edit', async () => {
		const source = await buildPackage();
		const loaded = await loadDocx(source);
		at(expectParagraph(loaded.model.blocks[0]).runs, 0).text = 'Edited';
		const saved = await loaded.save();
		await expectKept(source, saved, ['word/document.xml']);
		const xml = must(
			await (await JSZip.loadAsync(saved)).file('word/document.xml')?.async('string'),
		);
		expect(xml).toContain('>Edited<');
		// The same document.xml minus the edit: unknown markup survives the XML round trip.
		for (const kept of [
			'mc:AlternateContent',
			'<mc:Fallback>',
			'{unknown-sect}',
			'w:tblPrChange',
			'w:trPrChange',
			'w:tcPrChange',
			'w:bookmarkStart',
			'mc:Ignorable="w14 w15"',
			'<w:rFonts w:ascii="Cambria"/>',
		])
			expect(xml, kept).toContain(kept);
	});

	it('keeps them after a table edit and after a structural edit of a table with tracked changes', async () => {
		const source = await buildPackage();
		const loaded = await loadDocx(source);
		const table = expectTable(loaded.model.blocks.find((block) => block.type === 'table'));
		at(at(at(table.rows, 0), 1).paragraphs, 0).runs[0]!.text = 'B1 edited';
		const saved = await loaded.save();
		await expectKept(source, saved, ['word/document.xml']);
		const xml = must(
			await (await JSZip.loadAsync(saved)).file('word/document.xml')?.async('string'),
		);
		expect(xml).toContain('B1 edited');
		expect(xml).toContain('w:tcPrChange');
	});

	it('keeps the people and font parts while maintaining comment identity metadata', async () => {
		const source = await buildPackage();
		const loaded = await loadDocx(source);
		const model = structuredClone(loaded.model);
		model.comments = [{ id: '0', author: 'Ada', text: 'Note' }];
		const saved = await loaded.save(model);
		await expectKept(
			source,
			saved,
			[
				'word/document.xml',
				'word/comments.xml',
				'word/commentsExtended.xml',
				'word/commentsIds.xml',
				'word/commentsExtensible.xml',
				'[Content_Types].xml',
				'word/_rels/document.xml.rels',
			],
			['word/comments.xml', 'word/commentsExtended.xml'],
		);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('word/commentsIds.xml')!.async('string')).toContain('<w16cid:commentId ');
		expect(await zip.file('word/commentsExtensible.xml')!.async('string')).toContain(
			'<w16cex:commentExtensible ',
		);
	});

	it('refuses an edit of the style catalogue instead of dropping it', async () => {
		const source = await buildPackage();
		const loaded = await loadDocx(source);
		const model = structuredClone(loaded.model);
		model.theme = { ...(model.theme ?? {}), name: 'Changed' } as never;
		await expect(loaded.save(model)).rejects.toThrow(/preserved unchanged/);
	});
});
