import JSZip from 'jszip';
import { ImageBuilder, MediaBuilder, PptxHandler } from 'pptx-viewer-core';

/** The origin the deck links to; specs answer it with `page.route`. */
export const EXTERNAL_ORIGIN = 'https://external-content.test';

/** A 2x2 solid blue PNG. */
const BLUE_PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC';
/** Placeholder bytes: the spec only checks whether the browser asks for the file. */
const VIDEO_BYTES = 'data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE=';

/**
 * One slide with a picture and a video whose media relationships are all
 * linked (`TargetMode="External"`) to {@link EXTERNAL_ORIGIN} instead of
 * embedded, so opening the deck makes the browser contact that origin unless
 * external content is blocked.
 */
export async function externalContentDeck(): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].elements = [
		ImageBuilder.create(BLUE_PNG).position(60, 80).size(300, 220).build(),
		MediaBuilder.video(VIDEO_BYTES).position(420, 80).size(320, 180).build(),
	];
	data.slides[0].isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const rels = await zip.file(relsPath)!.async('string');
	let linked = 0;
	const rewritten = rels.replace(
		/(<Relationship\b[^>]*?)Target="\.\.\/media\/([^"]+)"/gu,
		(_match, before: string, file: string) => {
			linked += 1;
			return `${before}Target="${EXTERNAL_ORIGIN}/${file}" TargetMode="External"`;
		},
	);
	if (linked === 0) {
		throw new Error('externalContentDeck: no media relationship to link');
	}
	zip.file(relsPath, rewritten);
	for (const name of Object.keys(zip.files)) {
		if (name.startsWith('ppt/media/')) {
			zip.remove(name);
		}
	}
	return zip.generateAsync({ type: 'nodebuffer' });
}
