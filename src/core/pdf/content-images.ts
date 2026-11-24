/** Pictures a content stream draws: image XObjects and inline images, placed in reading space. */
import { apply, multiply, type Matrix } from './matrix';
import type { PdfDocument } from './document';
import { decodeImage, readImage, type ImageOptions } from './images';
import { inlineData, readInlineImage } from './inline-image';
import type { PdfLexer } from './lexer';
import type { PdfDict, PdfStream } from './objects';
import type { PdfPageImage } from './page-content';

export type Box = { x: number; y: number; width: number; height: number };

/** Where the unit square lands on the page. */
export function unitBox(ctm: Matrix, page: Matrix): Box {
	const toPage = multiply(ctm, page);
	const corners = [
		apply(toPage, 0, 0),
		apply(toPage, 1, 0),
		apply(toPage, 0, 1),
		apply(toPage, 1, 1),
	];
	const xs = corners.map((point) => point[0]);
	const ys = corners.map((point) => point[1]);
	const x = Math.min(...xs);
	const y = Math.min(...ys);
	return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
}

const failed = (box: Box): PdfPageImage => ({
	...box,
	pixelWidth: 0,
	pixelHeight: 0,
	format: 'unsupported',
	reason: 'it could not be decoded',
});

/** The picture of an image XObject drawn with `ctm`. */
export async function xobjectImage(
	document: PdfDocument,
	stream: PdfStream,
	box: Box,
	options: ImageOptions,
): Promise<PdfPageImage> {
	try {
		return await readImage(document, stream, box, options);
	} catch {
		return failed(box);
	}
}

/**
 * Reads the inline image after a `BI` operator (always, so the lexer moves past its data) and
 * decodes it when `decode` is set; undefined when it is malformed or not decoded.
 */
export async function inlineImage(
	document: PdfDocument,
	lexer: PdfLexer,
	resources: PdfDict,
	box: Box,
	options: ImageOptions & { decode: boolean },
): Promise<PdfPageImage | undefined> {
	const image = readInlineImage(lexer);
	if (!image || !options.decode) return undefined;
	try {
		return await decodeImage(
			document,
			{
				dict: image.dict,
				data: async () => inlineData(image, document.limits.maxStreamBytes),
				named: async (name) => (await document.dict(resources.get('ColorSpace')))?.get(name),
			},
			box,
			options,
		);
	} catch {
		return failed(box);
	}
}
