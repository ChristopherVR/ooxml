import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, tocEntries, type DocumentModel } from './index';
import { tocSwitches } from './toc-switches';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A package with a style catalog (Heading1, a Callout style at outline level 2) and `body` XML. */
async function packageWith(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout Box"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Sidebar"><w:name w:val="Sidebar"/><w:basedOn w:val="Normal"/></w:style></w:styles>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const para = (text: string, pPr = '') =>
	`<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

describe('TOC instruction switches', () => {
	it('reads \\o, \\u and \\t', () => {
		const all = tocSwitches(' TOC \\o "2-4" \\u \\t "Callout Box,3;Sidebar,1" ');
		expect(all).toMatchObject({ headings: true, outline: true, from: 2, to: 4 });
		expect([...all.styles]).toEqual([
			['callout box', 3],
			['sidebar', 1],
		]);
	});

	it('keeps every heading for a bare TOC but only the named styles for \\t alone', () => {
		expect(tocSwitches(' TOC ')).toMatchObject({ headings: true, outline: false });
		expect(tocSwitches(' TOC \\t "Sidebar,2" ')).toMatchObject({ headings: false });
		expect(tocSwitches(' TOC \\u ')).toMatchObject({ headings: false, outline: true });
	});

	it('ignores out-of-range levels and odd pairs in \\t', () => {
		expect(tocSwitches(' TOC \\t "A,0,B,10,C,x,D" ').styles.size).toBe(0);
	});
});

describe('TOC entries from custom styles and outline levels', () => {
	async function loadedBytes(): Promise<Uint8Array> {
		return packageWith(
			[
				para('Chapter', '<w:pStyle w:val="Heading1"/>'),
				para('Boxed', '<w:pStyle w:val="Callout"/>'),
				para('Side note', '<w:pStyle w:val="Sidebar"/>'),
				para('Promoted', '<w:outlineLvl w:val="2"/>'),
				para('Demoted', '<w:pStyle w:val="Heading1"/><w:outlineLvl w:val="9"/>'),
				para('Plain'),
			].join(''),
		);
	}
	async function model(): Promise<DocumentModel> {
		return (await loadDocx(await loadedBytes())).model;
	}
	const texts = (entries: { text: string; level: number }[]) =>
		entries.map((entry) => `${entry.level}:${entry.text}`);

	it('parses outline levels, a direct 9 meaning body text', async () => {
		const loaded = await model();
		const levels = loaded.blocks.map((block) =>
			block.type === 'paragraph' ? block.outlineLevel : undefined,
		);
		expect(levels).toEqual([undefined, undefined, undefined, 3, 0, undefined]);
		expect(loaded.paragraphStyles?.styles.Callout?.formatting.outlineLevel).toBe(2);
	});

	it('\\o lists headings only', async () => {
		const loaded = await model();
		expect(texts(tocEntries(loaded, ' TOC \\o "1-3" '))).toEqual(['1:Chapter', '1:Demoted']);
	});

	it('\\u adds paragraphs by outline level, honouring a direct body-text override', async () => {
		const loaded = await model();
		expect(texts(tocEntries(loaded, ' TOC \\o "1-3" \\u '))).toEqual([
			'1:Chapter',
			'2:Boxed',
			'3:Promoted',
			'1:Demoted',
		]);
		expect(texts(tocEntries(loaded, ' TOC \\u '))).toEqual(['1:Chapter', '2:Boxed', '3:Promoted']);
		expect(texts(tocEntries(loaded, ' TOC \\o "1-2" \\u '))).toEqual([
			'1:Chapter',
			'2:Boxed',
			'1:Demoted',
		]);
	});

	it('\\t maps a style by name or id to the given level', async () => {
		const loaded = await model();
		expect(texts(tocEntries(loaded, ' TOC \\t "Callout Box,2,Sidebar,3" '))).toEqual([
			'2:Boxed',
			'3:Side note',
		]);
		expect(texts(tocEntries(loaded, ' TOC \\o "1-1" \\t "sidebar,2" '))).toEqual([
			'1:Chapter',
			'2:Side note',
			'1:Demoted',
		]);
	});

	it('keeps the outline levels in the source XML when the document is saved', async () => {
		const handle = await loadDocx(await loadedBytes());
		const saved = await handle.save();
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:outlineLvl w:val="2"/>');
		expect(xml).toContain('<w:outlineLvl w:val="9"/>');
		const again = (await loadDocx(saved)).model;
		expect(again.blocks.map((b) => (b.type === 'paragraph' ? b.outlineLevel : undefined))).toEqual([
			undefined,
			undefined,
			undefined,
			3,
			0,
			undefined,
		]);
	});

	it('works on a model with no style catalog', () => {
		const bare = createDocument();
		bare.blocks = [
			{ type: 'paragraph', id: 'p', outlineLevel: 2, runs: [{ text: 'Only outline' }] },
		];
		expect(texts(tocEntries(bare, ' TOC \\u '))).toEqual(['2:Only outline']);
	});
});
