import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, pictureWatermarkSizePt, type Paragraph, type TextRun } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
		'base64',
	),
);
/** Shaped like the picture watermark Word writes: the t75 picture frame and washed-out imagedata. */
const PICTURE_RUN = `<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"/><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="${R}" id="WordPictureWatermark7" o:spid="_x0000_s2050" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:234pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:imagedata r:id="rId5" o:title="logo" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r>`;

async function fixture(headerBody: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${W}"><w:p>${headerBody}<w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/_rels/header1.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
	);
	zip.file('word/media/image1.png', PNG);
	return zip.generateAsync({ type: 'uint8array' });
}
const part = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';
const headerRuns = (model: Awaited<ReturnType<typeof loadDocx>>['model']) =>
	(model.sections![0]!.headers!.default!.blocks[0] as Paragraph).runs;
const watermarkRun = (runs: TextRun[]) => runs.find((run) => run.image?.pictureWatermark);

describe('picture watermarks', () => {
	it('reads Word picture watermark settings from the VML shape', async () => {
		const loaded = await loadDocx(await fixture(PICTURE_RUN));
		const image = watermarkRun(headerRuns(loaded.model))!.image!;
		expect(image).toMatchObject({
			unsupported: 'Watermark',
			partName: 'word/media/image1.png',
			contentType: 'image/png',
			widthPx: 0,
			heightPx: 0,
			pictureWatermark: {
				kind: 'picture',
				partName: 'word/media/image1.png',
				widthPt: 468,
				heightPt: 234,
				washout: true,
				title: 'logo',
			},
		});
		expect(image.watermark).toBeUndefined();
	});

	it('keeps the original shape verbatim when only the header text changes', async () => {
		const loaded = await loadDocx(await fixture(PICTURE_RUN));
		headerRuns(loaded.model).find((run) => run.text === 'Head')!.text = 'Header';
		const xml = await part(await loaded.save(), 'word/header1.xml');
		expect(xml).toContain('id="WordPictureWatermark7"');
		expect(xml).toContain('<w:t>Header</w:t>');
	});

	it('writes a new picture watermark with its relationship, then replaces and removes it', async () => {
		const loaded = await loadDocx(await fixture(''));
		const model = structuredClone(loaded.model);
		const runs = headerRuns(model);
		runs.unshift({
			text: '',
			image: {
				relId: '',
				partName: 'word/media/watermark1.png',
				contentType: 'image/png',
				widthPx: 400,
				heightPx: 200,
				unsupported: 'Watermark',
				pictureWatermark: {
					kind: 'picture',
					partName: 'word/media/watermark1.png',
					contentType: 'image/png',
					widthPt: 300,
					heightPt: 150,
					scale: 'auto',
					washout: true,
					title: 'seal',
				},
			},
		});
		const pending = new Map([
			['word/media/watermark1.png', { bytes: PNG, contentType: 'image/png' }],
		]);
		const saved = await loaded.save(model, pending);
		const xml = await part(saved, 'word/header1.xml');
		expect(xml).toMatch(/<v:shapetype id="_x0000_t75"[^>]*o:spt="75"/);
		expect(xml).toMatch(/<v:shape id="WordPictureWatermark\d+"[^>]*type="#_x0000_t75"/);
		expect(xml).toContain('width:300pt;height:150pt;z-index:-251657216');
		expect(xml).toContain('mso-position-horizontal:center;mso-position-horizontal-relative:margin');
		const rid =
			/<v:imagedata [^>]*r:id="(rId\d+)" o:title="seal" gain="19661f" blacklevel="22938f"\/>/.exec(
				xml,
			)?.[1];
		expect(rid).toBeTruthy();
		expect(xml).toMatch(/<w10:wrap [^>]*anchorx="margin" anchory="margin"[^>]*\/><\/v:shape>/);
		const rels = await part(saved, 'word/_rels/header1.xml.rels');
		expect(rels).toMatch(new RegExp(`Id="${rid}"[^>]*Target="media/watermark1.png"`));
		expect(await (await JSZip.loadAsync(saved)).file('word/media/watermark1.png')).toBeTruthy();

		const reopened = await loadDocx(saved);
		const spec = watermarkRun(headerRuns(reopened.model))!.image!.pictureWatermark!;
		expect(spec).toMatchObject({ widthPt: 300, heightPt: 150, washout: true, title: 'seal' });

		const plain = structuredClone(reopened.model);
		const plainRun = watermarkRun(headerRuns(plain))!;
		plainRun.image = {
			...plainRun.image!,
			pictureWatermark: { ...spec, washout: false },
		};
		headerRuns(plain).splice(headerRuns(plain).indexOf(plainRun), 1, { ...plainRun });
		const replaced = await part(await reopened.save(plain), 'word/header1.xml');
		expect(replaced).not.toContain('gain=');

		const removed = structuredClone(reopened.model);
		const paragraph = removed.sections![0]!.headers!.default!.blocks[0] as Paragraph;
		paragraph.runs = paragraph.runs.filter((run) => !run.image?.pictureWatermark);
		const cleared = await part(await reopened.save(removed), 'word/header1.xml');
		expect(cleared).not.toContain('WordPictureWatermark');
		expect(cleared).toContain('<w:t>Head</w:t>');
	});

	it('sizes Auto to fit the text area and percentages from the natural size', () => {
		const natural = { widthPt: 150, heightPt: 75 };
		expect(pictureWatermarkSizePt(natural, 'auto', { widthPt: 468, heightPt: 648 })).toEqual({
			widthPt: 468,
			heightPt: 234,
		});
		expect(pictureWatermarkSizePt(natural, 200, { widthPt: 468, heightPt: 648 })).toEqual({
			widthPt: 300,
			heightPt: 150,
		});
	});
});
