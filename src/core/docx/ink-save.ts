// Saving Word ink. Ink drawn in this session gets a new InkML part (`word/ink/inkN.xml`) with its
// content type; ink whose strokes changed (erased, merged, moved by the lasso, resized) is written
// back over its own part, keeping the nodes the model does not hold. Ink nobody touched keeps its
// part byte for byte, Word's own trace encoding included. The InkML writer loads on first use.
import type JSZip from 'jszip';
import type { InkStroke } from '../drawingml/inkml/types';
import { documentDrawingBlocks } from './chart';
import { inksIn, type DocxInk } from './ink';
import { forEachParagraph } from './parse-warnings';
import type { DocumentModel, InlineImage } from './model';
import { ensureContentTypeOverride } from './zip-parts';

const INKML_CONTENT_TYPE = 'application/inkml+xml';

/** Writes the part of ink drawn in this session and returns its name. */
export async function writeNewInkPart(
	zip: JSZip,
	drawn: { ink: DocxInk; widthPx: number; heightPx: number },
): Promise<string> {
	const { writeInkPart } = await import('#inkml-runtime');
	const strokes = await inkStrokesForBox(drawn.ink, drawn.widthPx, drawn.heightPx);
	let index = 1;
	while (zip.file(`word/ink/ink${index}.xml`)) index++;
	const partName = `word/ink/ink${index}.xml`;
	zip.file(partName, writeInkPart(strokes));
	await ensureContentTypeOverride(zip, partName, INKML_CONTENT_TYPE);
	return partName;
}

/**
 * The strokes of an ink drawing in its saved box: drawn in a space of the box's size, the part
 * reads back where it was drawn; after a resize they are stretched to the box, as Word draws them.
 */
export async function inkStrokesForBox(
	ink: DocxInk,
	widthPx: number,
	heightPx: number,
): Promise<InkStroke[]> {
	if (ink.width === widthPx && ink.height === heightPx) return ink.strokes;
	const { transformInkStrokes } = await import('#inkml-runtime');
	return transformInkStrokes(
		ink.strokes,
		0,
		0,
		ink.width > 0 ? widthPx / ink.width : 1,
		ink.height > 0 ? heightPx / ink.height : 1,
	);
}

function inkImages(model: DocumentModel): InlineImage[] {
	const images: InlineImage[] = [];
	forEachParagraph(documentDrawingBlocks(model), (paragraph) => {
		for (const run of paragraph.runs) if (run.image?.ink?.partName) images.push(run.image);
	});
	return images;
}

/** Rewrites every loaded ink part whose strokes or box changed since `base`. */
export async function applyInkPartEdits(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel | undefined,
): Promise<void> {
	if (!base) return;
	const loaded = new Map<string, string>();
	for (const ink of inksIn(documentDrawingBlocks(base)))
		if (ink.partName) loaded.set(ink.partName, JSON.stringify(ink));
	const images = inkImages(model).filter((image) => {
		const ink = image.ink!;
		const before = loaded.get(ink.partName!);
		return (
			before !== undefined &&
			(before !== JSON.stringify(ink) ||
				ink.width !== image.widthPx ||
				ink.height !== image.heightPx)
		);
	});
	if (!images.length) return;
	const { updatedInkPart } = await import('#inkml-runtime');
	const saved = new Set<string>();
	for (const image of images) {
		const ink = image.ink!;
		if (saved.has(ink.partName!)) continue;
		saved.add(ink.partName!);
		const box = { width: image.widthPx, height: image.heightPx };
		const strokes = await inkStrokesForBox(ink, box.width, box.height);
		const source = await zip.file(ink.partName!)?.async('string');
		const xml = updatedInkPart(source, strokes, box);
		if (xml !== undefined) zip.file(ink.partName!, xml);
	}
}
