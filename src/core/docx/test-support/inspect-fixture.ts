import JSZip from 'jszip';

// A package shaped like the one Word 16 saves for a document with every kind of information
// Inspect Document reports; the names, dates and paths are invented.
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

export const AUTHOR = 'Alice Example';
export const REVIEWER = 'Bob Reviewer';
export const TEMPLATE_PATH = 'file:///C:/Templates/Private/Letter.dotx';

const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const override = (part: string, type: string) =>
	`<Override PartName="/${part}" ContentType="${type}"/>`;

export async function inspectFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`${decl}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[
			override('word/document.xml', `${CT}.document.main+xml`),
			override(
				'customXml/itemProps1.xml',
				'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
			),
			override('word/settings.xml', `${CT}.settings+xml`),
			override('word/comments.xml', `${CT}.comments+xml`),
			override('word/commentsExtensible.xml', `${CT}.commentsExtensible+xml`),
			override('word/commentsIds.xml', `${CT}.commentsIds+xml`),
			override('word/header1.xml', `${CT}.header+xml`),
			override('word/footer1.xml', `${CT}.footer+xml`),
			override('word/people.xml', `${CT}.people+xml`),
			override('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
			override(
				'docProps/app.xml',
				'application/vnd.openxmlformats-officedocument.extended-properties+xml',
			),
			override(
				'docProps/custom.xml',
				'application/vnd.openxmlformats-officedocument.custom-properties+xml',
			),
		].join('')}</Types>`,
	);
	zip.file(
		'_rels/.rels',
		`${decl}<Relationships xmlns="${PKG}"><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/></Relationships>`,
	);
	zip.file(
		'docProps/core.xml',
		`${decl}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Quarterly plan</dc:title><dc:creator>${AUTHOR}</dc:creator><cp:keywords>plan</cp:keywords><cp:lastModifiedBy>${REVIEWER}</cp:lastModifiedBy><cp:revision>3</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-03T03:04:00Z</dcterms:modified><cp:category>Plans</cp:category></cp:coreProperties>`,
	);
	zip.file(
		'docProps/app.xml',
		`${decl}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Letter.dotx</Template><TotalTime>4</TotalTime><Application>Microsoft Office Word</Application><Manager>Carol Manager</Manager><Company>Example Co</Company><AppVersion>16.0000</AppVersion></Properties>`,
	);
	zip.file(
		'docProps/custom.xml',
		`${decl}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Project"><vt:lpwstr>Apollo</vt:lpwstr></property></Properties>`,
	);
	zip.file('customXml/item1.xml', '<data xmlns="urn:example"><v>1</v></data>');
	zip.file(
		'customXml/itemProps1.xml',
		`${decl}<ds:datastoreItem ds:itemID="{11111111-2222-3333-4444-555555555555}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>`,
	);
	zip.file(
		'customXml/_rels/item1.xml.rels',
		`${decl}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`${decl}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/customXml" Target="../customXml/item1.xml"/><Relationship Id="rId2" Type="${REL}/settings" Target="settings.xml"/><Relationship Id="rId3" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/><Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/><Relationship Id="rId6" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rId7" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rId8" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`${decl}<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w16du="${W16DU}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Visible paragraph one.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Shown </w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden words</w:t></w:r></w:p><w:p><w:ins w:id="1" w:author="${REVIEWER}" w:date="2026-01-02T05:00:00Z" w16du:dateUtc="2026-01-02T05:00:00Z"><w:r><w:t>Tracked insertion.</w:t></w:r></w:ins></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId6"/><w:footerReference w:type="default" r:id="rId7"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/comments.xml',
		`${decl}<w:comments xmlns:w="${W}" xmlns:w14="${W14}"><w:comment w:id="0" w:author="${AUTHOR}" w:date="2026-01-02T04:00:00Z" w:initials="AE"><w:p w14:paraId="1A2B3C4D"><w:r><w:t>A comment</w:t></w:r></w:p></w:comment></w:comments>`,
	);
	zip.file(
		'word/commentsIds.xml',
		`${decl}<w16cid:commentsIds xmlns:w16cid="${W16CID}"><w16cid:commentId w16cid:paraId="1A2B3C4D" w16cid:durableId="5E6F7A8B"/></w16cid:commentsIds>`,
	);
	zip.file(
		'word/commentsExtensible.xml',
		`${decl}<w16cex:commentsExtensible xmlns:w16cex="${W16CEX}"><w16cex:commentExtensible w16cex:durableId="5E6F7A8B" w16cex:dateUtc="2026-01-02T04:00:00Z"/></w16cex:commentsExtensible>`,
	);
	zip.file(
		'word/people.xml',
		`${decl}<w15:people xmlns:w15="${W15}"><w15:person w15:author="${AUTHOR}"><w15:presenceInfo w15:providerId="None" w15:userId="${AUTHOR}"/></w15:person></w15:people>`,
	);
	zip.file(
		'word/header1.xml',
		`${decl}<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/footer1.xml',
		`${decl}<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>`,
	);
	zip.file(
		'word/settings.xml',
		`${decl}<w:settings xmlns:w="${W}" xmlns:r="${R}" xmlns:w15="${W15}"><w:zoom w:percent="100"/><w:attachedTemplate r:id="rId1"/><w:defaultTabStop w:val="720"/><w15:docId w15:val="{82AFE0D4-4879-40C8-B5F6-9CB108FC40E5}"/></w:settings>`,
	);
	zip.file(
		'word/_rels/settings.xml.rels',
		`${decl}<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/attachedTemplate" Target="${TEMPLATE_PATH}" TargetMode="External"/></Relationships>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Every XML part of a saved package, by name. */
export async function xmlParts(bytes: Uint8Array): Promise<Map<string, string>> {
	const zip = await JSZip.loadAsync(bytes);
	const parts = new Map<string, string>();
	for (const file of Object.values(zip.files))
		if (!file.dir && /\.(xml|rels)$/.test(file.name))
			parts.set(file.name, await file.async('string'));
	return parts;
}
