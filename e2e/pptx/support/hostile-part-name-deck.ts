import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

/** Marks anything the deck's slide part name manages to add to the page. */
export const INJECTED_ATTRIBUTE = 'data-injected';

const SLIDE_PART = 'ppt/slides/slide1.xml';
const FILE_NAME = `s</style><i ${INJECTED_ATTRIBUTE}="1">.xml`;
const FILE_NAME_IN_XML = FILE_NAME.replaceAll('<', '&lt;')
	.replaceAll('>', '&gt;')
	.replaceAll('"', '&quot;');

/** A clipped preset with a fill, a glow and text: its hit area carries a per-element style rule. */
const GLOWING_TRIANGLE =
	'<p:sp><p:nvSpPr><p:cNvPr id="900" name="Triangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C83232"/></a:solidFill><a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Glow</a:t></a:r></a:p></p:txBody></p:sp>';

async function replaceInPart(zip: JSZip, path: string, from: string, to: string): Promise<void> {
	const xml = await zip.file(path)?.async('string');
	if (!xml?.includes(from)) {
		throw new Error(`${path} does not contain ${from}`);
	}
	zip.file(path, xml.replace(from, to));
}

/** One slide stored under a part name that holds markup, as a crafted package can do. */
export async function hostilePartNameDeck(): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();

	const slide = await zip.file(SLIDE_PART)?.async('string');
	if (!slide?.includes('</p:spTree>')) {
		throw new Error(`${SLIDE_PART} has no shape tree`);
	}
	zip.remove(SLIDE_PART);
	zip.file(
		`ppt/slides/${FILE_NAME}`,
		slide.replace('</p:spTree>', `${GLOWING_TRIANGLE}</p:spTree>`),
	);
	await replaceInPart(
		zip,
		'ppt/_rels/presentation.xml.rels',
		'slides/slide1.xml',
		`slides/${FILE_NAME_IN_XML}`,
	);
	await replaceInPart(
		zip,
		'[Content_Types].xml',
		`/${SLIDE_PART}`,
		`/ppt/slides/${FILE_NAME_IN_XML}`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}
