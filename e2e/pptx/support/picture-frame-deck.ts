import JSZip from 'jszip';
import { ImageBuilder, PptxHandler } from 'pptx-viewer-core';

/** A 2x2 solid blue PNG. */
const BLUE_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC';

/**
 * One picture carrying what a Picture Styles pick writes: an 8pt red frame
 * (`a:ln`, centred on the edge) and an outer shadow (`a:outerShdw`). The two
 * are written into the saved slide XML so the deck is exactly what PowerPoint
 * produces, independent of the editor's own style writer.
 */
export async function pictureFrameDeck(): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].elements = [
		ImageBuilder.create(BLUE_PNG).position(200, 120).size(300, 220).build(),
	];
	data.slides[0].isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	const frame =
		'<a:ln w="101600"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>' +
		'<a:effectLst><a:outerShdw blurRad="50800" dist="101600" dir="2700000" algn="tl" rotWithShape="0">' +
		'<a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>';
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const framed = xml.replace(/(<p:pic>[\s\S]*?<a:prstGeom[\s\S]*?<\/a:prstGeom>)/u, `$1${frame}`);
	if (framed === xml) {
		throw new Error('pictureFrameDeck: no picture geometry to frame');
	}
	zip.file('ppt/slides/slide1.xml', framed);
	return zip.generateAsync({ type: 'nodebuffer' });
}
