import JSZip from 'jszip';
import { ImageBuilder, MediaBuilder, PptxHandler } from 'pptx-viewer-core';
import type { Model3DPptxElement } from 'pptx-viewer-core';

/** The origin the deck links to; specs answer it with `page.route`. */
export const EXTERNAL_ORIGIN = 'https://external-content.test';

/** The linked pictures each kind of element references, relative to {@link EXTERNAL_ORIGIN}. */
export const LINKED_PICTURES = {
	shapeFill: '/fill.png',
	pictureBullet: '/bullet.png',
} as const;

/** How many elements the deck's slide has, so a spec can wait for all of them to render. */
export const EXTERNAL_CONTENT_DECK_ELEMENTS = 5;

/** A 2x2 solid blue PNG. */
const BLUE_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC';
/** Placeholder bytes: the spec only checks whether the browser asks for the file. */
const VIDEO_BYTES = 'data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE=';
const GLB_BYTES = 'data:model/gltf-binary;base64,Z2xURgIAAAAMAAAA';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const LINKED_FILL_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="571500" y="3429000"/><a:ext cx="1371600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;
const LINKED_BULLET_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="902" name="Linked bullet"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2286000" y="3429000"/><a:ext cx="3657600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdLinkedBullet"/></a:buBlip></a:pPr><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Linked bullet</a:t></a:r></a:p></p:txBody></p:sp>`;

/**
 * One slide whose pictures and media are all linked (`TargetMode="External"`)
 * to {@link EXTERNAL_ORIGIN} instead of embedded: a picture, a video, a 3D
 * model's poster, a shape with a picture fill and a picture bullet. Opening
 * the deck makes the browser contact that origin unless external content is
 * blocked. The 3D model itself stays embedded.
 */
export async function externalContentDeck(): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].elements = [
		ImageBuilder.create(BLUE_PNG).position(60, 80).size(300, 220).build(),
		MediaBuilder.video(VIDEO_BYTES).position(420, 80).size(320, 180).build(),
		{
			id: 'linked-model',
			type: 'model3d',
			x: 780,
			y: 80,
			width: 160,
			height: 120,
			modelData: GLB_BYTES,
			modelMimeType: 'model/gltf-binary',
			posterImage: BLUE_PNG,
		} satisfies Model3DPptxElement,
	];
	data.slides[0].isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	const slidePath = 'ppt/slides/slide1.xml';
	const slideXml = await zip.file(slidePath)!.async('string');
	zip.file(
		slidePath,
		slideXml.replace('</p:spTree>', `${LINKED_FILL_SHAPE}${LINKED_BULLET_SHAPE}</p:spTree>`),
	);
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const rels = await zip.file(relsPath)!.async('string');
	let linked = 0;
	const rewritten = rels.replace(
		/(<Relationship\b[^>]*?)Target="\.\.\/media\/([^"]+\.(?:png|mp4))"/gu,
		(_match, before: string, file: string) => {
			linked += 1;
			return `${before}Target="${EXTERNAL_ORIGIN}/${file}" TargetMode="External"`;
		},
	);
	if (linked === 0) {
		throw new Error('externalContentDeck: no media relationship to link');
	}
	zip.file(
		relsPath,
		rewritten.replace(
			'</Relationships>',
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${EXTERNAL_ORIGIN}${LINKED_PICTURES.shapeFill}" TargetMode="External"/>` +
				`<Relationship Id="rIdLinkedBullet" Type="${IMAGE_REL}" Target="${EXTERNAL_ORIGIN}${LINKED_PICTURES.pictureBullet}" TargetMode="External"/></Relationships>`,
		),
	);
	for (const name of Object.keys(zip.files)) {
		if (/^ppt\/media\/.+\.(?:png|mp4)$/u.test(name)) {
			zip.remove(name);
		}
	}
	return zip.generateAsync({ type: 'nodebuffer' });
}
