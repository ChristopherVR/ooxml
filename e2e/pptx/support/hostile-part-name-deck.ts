import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

/** Marks anything the deck's slide part name manages to add to the page. */
export const INJECTED_ATTRIBUTE = 'data-injected';

const SLIDE_PART = 'ppt/slides/slide1.xml';
const SLIDE_RELATIONSHIPS = 'ppt/slides/_rels/slide1.xml.rels';
const HOSTILE_FILE_NAME = `s</style><i ${INJECTED_ATTRIBUTE}="1">.xml`;

/** A clipped preset with a fill, a glow and text: its hit area carries a per-element style rule. */
const GLOWING_TRIANGLE =
	'<p:sp><p:nvSpPr><p:cNvPr id="900" name="Triangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C83232"/></a:solidFill><a:effectLst><a:glow rad="63500"><a:srgbClr val="FF0000"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Glow</a:t></a:r></a:p></p:txBody></p:sp>';

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;');
}

async function replaceInPart(zip: JSZip, path: string, from: string, to: string): Promise<void> {
	const xml = await zip.file(path)?.async('string');
	if (!xml?.includes(from)) {
		throw new Error(`${path} does not contain ${from}`);
	}
	zip.file(path, xml.replace(from, to));
}

/** One slide holding `shapes`, stored as `ppt/slides/<fileName>`. */
export async function deckWithSlidePart(fileName: string, shapes: string): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();

	const slide = await zip.file(SLIDE_PART)?.async('string');
	if (!slide?.includes('</p:spTree>')) {
		throw new Error(`${SLIDE_PART} has no shape tree`);
	}
	const nameInXml = escapeXml(fileName);
	const relationships = await zip.file(SLIDE_RELATIONSHIPS)?.async('uint8array');
	if (!relationships) {
		throw new Error(`${SLIDE_RELATIONSHIPS} is missing`);
	}
	zip.remove(SLIDE_PART);
	zip.remove(SLIDE_RELATIONSHIPS);
	zip.file(`ppt/slides/${fileName}`, slide.replace('</p:spTree>', `${shapes}</p:spTree>`));
	zip.file(`ppt/slides/_rels/${fileName}.rels`, relationships);
	await replaceInPart(
		zip,
		'ppt/_rels/presentation.xml.rels',
		'slides/slide1.xml',
		`slides/${nameInXml}`,
	);
	await replaceInPart(zip, '[Content_Types].xml', `/${SLIDE_PART}`, `/ppt/slides/${nameInXml}`);
	return zip.generateAsync({ type: 'nodebuffer' });
}

/** One slide stored under a part name that holds markup, as a crafted package can do. */
export function hostilePartNameDeck(): Promise<Buffer> {
	return deckWithSlidePart(HOSTILE_FILE_NAME, GLOWING_TRIANGLE);
}
