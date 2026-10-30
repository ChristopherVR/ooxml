import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse';
import { createListDefinition } from './numbering-editing';
import { computeListLabels } from './numbering-format';
import { pictureBulletSizePt } from './numbering-picture';
import type { Paragraph } from './model';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PNG = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3]);

async function pictureBulletDocx(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<?xml version="1.0"?><w:numbering xmlns:w="${W}" xmlns:r="${R}" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:9pt;height:12pt"><v:imagedata r:id="rId1" o:title="bullet"/></v:shape></w:pict></w:numPicBullet><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlPicBulletId w:val="0"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	zip.file(
		'word/_rels/numbering.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
	);
	zip.file('word/media/image1.png', PNG);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('picture bullets', () => {
	it('resolves numPicBullet through the numbering relationships and draws it as the label', async () => {
		const loaded = await loadDocx(await pictureBulletDocx());
		const catalog = loaded.model.numberingCatalog!;
		expect(catalog.pictureBullets?.['0']).toEqual({
			id: '0',
			partName: 'word/media/image1.png',
			contentType: 'image/png',
			widthPt: 9,
			heightPt: 12,
		});
		expect(catalog.abstractNums['0']?.levels[0]?.pictureBulletId).toBe('0');
		expect(catalog.warnings.join()).not.toContain('picture bullet');
		expect(loaded.media?.get('word/media/image1.png')).toEqual(PNG);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const label = computeListLabels(loaded.model).get(paragraph.id);
		expect(label?.picture?.partName).toBe('word/media/image1.png');
	});

	it('sizes to the font unless the shape authors a size', () => {
		expect(pictureBulletSizePt({}, 11)).toEqual({ widthPt: 11, heightPt: 11 });
		expect(pictureBulletSizePt({ widthPt: 9, heightPt: 12 }, 11)).toEqual({
			widthPt: 9,
			heightPt: 12,
		});
	});

	it('round-trips unchanged and writes lvlPicBulletId for a new list reusing the picture', async () => {
		const bytes = await pictureBulletDocx();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		const { catalog, numId } = createListDefinition(loaded.model.numberingCatalog, [
			{ level: 0, start: 1, numFmt: 'bullet', lvlText: '', pictureBulletId: '0' },
		]);
		const paragraph: Paragraph = {
			type: 'paragraph',
			id: 'extra',
			runs: [{ text: 'More' }],
			numbering: { numId, level: 0 },
		};
		const output = await JSZip.loadAsync(
			await loaded.save({
				...loaded.model,
				blocks: [...loaded.model.blocks, paragraph],
				numberingCatalog: catalog,
			}),
		);
		const xml = (await output.file('word/numbering.xml')?.async('string')) ?? '';
		expect(xml).toContain('<w:numPicBullet');
		expect(xml.match(/w:lvlPicBulletId/g)).toHaveLength(2);
		expect(await output.file('word/media/image1.png')?.async('uint8array')).toEqual(PNG);
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		const labels = computeListLabels(reopened.model);
		const reopenedExtra = reopened.model.blocks[1] as Paragraph;
		expect(labels.get(reopenedExtra.id)?.picture?.id).toBe('0');
	});

	it('rejects a new level that references an undefined picture bullet', async () => {
		const loaded = await loadDocx(await pictureBulletDocx());
		const { catalog, numId } = createListDefinition(loaded.model.numberingCatalog, [
			{ level: 0, start: 1, numFmt: 'bullet', lvlText: '*', pictureBulletId: '9' },
		]);
		const paragraph: Paragraph = {
			type: 'paragraph',
			id: 'bad',
			runs: [{ text: 'x' }],
			numbering: { numId, level: 0 },
		};
		await expect(
			loaded.save({ ...loaded.model, blocks: [paragraph], numberingCatalog: catalog }),
		).rejects.toThrow(/picture bullet "9"/);
	});
});
