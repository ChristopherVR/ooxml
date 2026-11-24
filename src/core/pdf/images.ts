/**
 * PDF images (XObjects and inline images) as pictures a document can hold: a DCT (JPEG) image
 * passes through as it is (a CMYK one with its Adobe marker set to keep the PDF's colours);
 * samples (1 to 16 bits, Flate, LZW, run length or CCITT fax coded) in any colour space this
 * reader knows become a PNG, with the /Decode array applied and the soft mask, stencil mask or
 * colour key as its alpha. JPEG 2000 and JBIG2 images are reported, not decoded.
 */
import { encodePng } from '../drawingml/raster-encode';
import { jpegComponents } from '../drawingml/raster-decode';
import { decodeCcitt } from './ccitt';
import type { PdfDocument } from './document';
import { colorSpace, gray, type ColorSpace } from './image-color';
import { cmykJpegForDocument } from './jpeg-cmyk';
import { isNumber, PdfStream, type PdfDict, type PdfValue } from './objects';
import type { PdfPageImage } from './page-content';

const MAX_PIXELS = 40_000_000;

type Decoded = { bytes: Uint8Array; imageFilter?: { name: string; parms: PdfDict | undefined } };

/** Where an image's dictionary and data come from (an XObject stream or an inline image). */
export interface ImageSource {
	dict: PdfDict;
	data(): Promise<Decoded>;
	/** Colour spaces named in the resources (inline images may use them). */
	named?: (name: string) => Promise<PdfValue | undefined>;
}

export interface ImageOptions {
	/** The fill colour a stencil mask paints, `#rrggbb`. */
	fill?: string;
}

export const streamSource = (document: PdfDocument, stream: PdfStream): ImageSource => ({
	dict: stream.dict,
	data: () => document.streamData(stream),
});

const UNSUPPORTED: Record<string, string> = {
	JPXDecode: 'its JPEG 2000 compression is not supported',
	JBIG2Decode: 'its JBIG2 compression is not supported',
};

/** The samples of decoded image data (CCITT fax decoded here), or why there are none. */
function samples(
	document: PdfDocument,
	decoded: Decoded,
	width: number,
	height: number,
): { bytes: Uint8Array; fax: boolean } | string {
	const filter = decoded.imageFilter;
	if (!filter) return { bytes: decoded.bytes, fax: false };
	if (filter.name === 'CCITTFaxDecode') {
		const parm = (key: string) => filter.parms?.get(key);
		const number = (key: string, fallback: number) => {
			const value = parm(key);
			return isNumber(value) ? value : fallback;
		};
		const bytes = decodeCcitt(
			decoded.bytes,
			{
				k: number('K', 0),
				columns: number('Columns', width),
				rows: number('Rows', 0) || height,
				encodedByteAlign: parm('EncodedByteAlign') === true,
				blackIs1: parm('BlackIs1') === true,
			},
			document.limits.maxStreamBytes,
		);
		return { bytes, fax: true };
	}
	return (
		UNSUPPORTED[filter.name] ??
		`its ${filter.name.replace(/Decode$/, '')} compression is not supported`
	);
}

/** Per component, a table from the raw sample to its 0..255 value (or palette index). */
function lookups(
	decode: PdfValue | undefined,
	components: number,
	bpc: number,
	hival: number | undefined,
): Uint8Array[] {
	const max = bpc === 16 ? 255 : (1 << bpc) - 1;
	const numbers = Array.isArray(decode) && decode.every(isNumber) ? (decode as number[]) : [];
	return Array.from({ length: components }, (_, c) => {
		const low = numbers.length >= 2 * components ? numbers[2 * c]! : 0;
		const high =
			numbers.length >= 2 * components ? numbers[2 * c + 1]! : hival === undefined ? 1 : max;
		const table = new Uint8Array(max + 1);
		for (let raw = 0; raw <= max; raw++) {
			const value = low + (raw * (high - low)) / max;
			table[raw] =
				hival === undefined
					? Math.min(255, Math.max(0, Math.round(value * 255)))
					: Math.min(hival, Math.max(0, Math.round(value)));
		}
		return table;
	});
}

/** Reads raw sample `component` of pixel (x, y) from packed rows (16-bit samples by their high byte). */
function reader(data: Uint8Array, width: number, components: number, bpc: number) {
	const stride = Math.ceil((width * components * bpc) / 8);
	return (x: number, y: number, component: number): number => {
		const index = x * components + component;
		if (bpc === 8) return data[y * stride + index] ?? 0;
		if (bpc === 16) return data[y * stride + index * 2] ?? 0;
		const bit = index * bpc;
		const byte = data[y * stride + (bit >> 3)] ?? 0;
		return (byte >> (8 - bpc - (bit & 7))) & ((1 << bpc) - 1);
	};
}

/** A mask image as an alpha plane at the image's size (nearest sample), or undefined. */
async function maskPlane(
	document: PdfDocument,
	mask: PdfStream,
	width: number,
	height: number,
	stencil: boolean,
): Promise<Uint8Array | undefined> {
	const get = (key: string) => document.resolve(mask.dict.get(key));
	const mw = await get('Width');
	const mh = await get('Height');
	if (!isNumber(mw) || !isNumber(mh) || mw < 1 || mh < 1 || mw * mh > MAX_PIXELS) return undefined;
	const raw = samples(document, await document.streamData(mask), mw, mh);
	if (typeof raw === 'string') return undefined;
	const bpcValue = await get('BitsPerComponent');
	const bpc = stencil || raw.fax ? 1 : isNumber(bpcValue) ? bpcValue : 8;
	if (![1, 2, 4, 8, 16].includes(bpc)) return undefined;
	const table = lookups(await get('Decode'), 1, bpc, undefined)[0]!;
	const read = reader(raw.bytes, mw, 1, bpc);
	const alpha = new Uint8Array(width * height);
	for (let y = 0; y < height; y++) {
		const my = Math.min(mh - 1, Math.floor((y * mh) / height));
		for (let x = 0; x < width; x++) {
			const value = table[read(Math.min(mw - 1, Math.floor((x * mw) / width)), my, 0)]!;
			// A stencil paints where its sample is 0 (after /Decode); a soft mask is the alpha.
			alpha[y * width + x] = stencil ? (value === 0 ? 255 : 0) : value;
		}
	}
	return alpha;
}

const parseHex = (color: string | undefined): [number, number, number] => {
	const match = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color ?? '');
	return match
		? ([1, 2, 3].map((index) => parseInt(match[index]!, 16)) as [number, number, number])
		: [0, 0, 0];
};

/** The picture an image XObject holds. */
export function readImage(
	document: PdfDocument,
	stream: PdfStream,
	box: { x: number; y: number; width: number; height: number },
	options: ImageOptions = {},
): Promise<PdfPageImage> {
	return decodeImage(document, streamSource(document, stream), box, options);
}

/** The picture of an image source. */
export async function decodeImage(
	document: PdfDocument,
	source: ImageSource,
	box: { x: number; y: number; width: number; height: number },
	options: ImageOptions = {},
): Promise<PdfPageImage> {
	const get = (key: string) => document.resolve(source.dict.get(key));
	const width = await get('Width');
	const height = await get('Height');
	const base = {
		...box,
		pixelWidth: isNumber(width) ? width : 0,
		pixelHeight: isNumber(height) ? height : 0,
	};
	const unsupported = (reason: string): PdfPageImage => ({
		...base,
		format: 'unsupported',
		reason,
	});
	if (!isNumber(width) || !isNumber(height) || width < 1 || height < 1)
		return unsupported('it has no size');
	if (width * height > MAX_PIXELS) return unsupported('it is larger than the size limit');
	const decoded = await source.data();
	const decode = await get('Decode');
	if (decoded.imageFilter?.name === 'DCTDecode') {
		const components = jpegComponents(decoded.bytes);
		if (components === 1 || components === 3)
			return { ...base, format: 'jpeg', data: decoded.bytes };
		if (components !== 4) return unsupported('its JPEG data could not be read');
		const inverted = Array.isArray(decode) && decode[0] === 1 && decode[1] === 0;
		const data = cmykJpegForDocument(decoded.bytes, inverted);
		return data
			? { ...base, format: 'jpeg', data }
			: unsupported('it is a YCCK JPEG whose colours would need decoding');
	}
	const raw = samples(document, decoded, width, height);
	if (typeof raw === 'string') return unsupported(raw);
	const imageMask = (await get('ImageMask')) === true;
	const bpcValue = await get('BitsPerComponent');
	const bpc = imageMask || raw.fax ? 1 : isNumber(bpcValue) ? bpcValue : 8;
	if (![1, 2, 4, 8, 16].includes(bpc)) return unsupported('its sample size is not supported');
	const space: ColorSpace | undefined = imageMask
		? gray
		: await colorSpace(document, source.dict.get('ColorSpace'), source.named);
	if (!space) return unsupported('its colour space is not supported');
	const tables = lookups(decode, space.components, bpc, space.hival);
	const read = reader(raw.bytes, width, space.components, bpc);
	const out = new Uint8Array(width * height * 3);
	const pixel: number[] = [];
	const colorKey = await get('Mask');
	const keyRanges =
		Array.isArray(colorKey) && colorKey.length >= 2 * space.components && colorKey.every(isNumber)
			? (colorKey as number[])
			: undefined;
	let alpha: Uint8Array | undefined =
		keyRanges || imageMask ? new Uint8Array(width * height) : undefined;
	const fill = parseHex(options.fill);
	for (let y = 0; y < height; y++)
		for (let x = 0; x < width; x++) {
			const index = y * width + x;
			let keyed = !!keyRanges;
			for (let c = 0; c < space.components; c++) {
				const sample = read(x, y, c);
				if (keyed && (sample < keyRanges![2 * c]! || sample > keyRanges![2 * c + 1]!))
					keyed = false;
				pixel[c] = tables[c]![sample]!;
			}
			if (imageMask) {
				out.set(fill, index * 3);
				alpha![index] = pixel[0] === 0 ? 255 : 0;
				continue;
			}
			space.toRgb(pixel, out, index * 3);
			if (keyRanges) alpha![index] = keyed ? 0 : 255;
		}
	const mask = await get('SMask');
	const stencil = await get('Mask');
	if (!imageMask && mask instanceof PdfStream) {
		alpha = (await maskPlane(document, mask, width, height, false)) ?? alpha;
		const matte = await document.resolve(mask.dict.get('Matte'));
		if (alpha && Array.isArray(matte) && matte.length === space.components && matte.every(isNumber))
			unpremultiply(out, alpha, matte as number[], space);
	} else if (!imageMask && stencil instanceof PdfStream)
		alpha = (await maskPlane(document, stencil, width, height, true)) ?? alpha;
	return { ...base, format: 'png', data: encodePng(width, height, out, alpha) };
}

/** Undoes the premultiplication of a soft mask's /Matte colour: c = m + (c' - m) / a. */
function unpremultiply(out: Uint8Array, alpha: Uint8Array, matte: number[], space: ColorSpace) {
	const color = new Uint8Array(3);
	space.toRgb(
		matte.map((value) => Math.round(Math.min(1, Math.max(0, value)) * 255)),
		color,
		0,
	);
	for (let index = 0; index < alpha.length; index++) {
		const a = alpha[index]! / 255;
		if (a <= 0) continue;
		for (let c = 0; c < 3; c++) {
			const at = index * 3 + c;
			out[at] = Math.min(255, Math.max(0, Math.round(color[c]! + (out[at]! - color[c]!) / a)));
		}
	}
}
