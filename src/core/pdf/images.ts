/**
 * PDF image XObjects as pictures a document can hold: a DCT (JPEG) image passes through as it is;
 * 8-bit (and 1-bit) samples in DeviceGray, DeviceRGB, DeviceCMYK, ICCBased or Indexed colour
 * become a PNG, with the soft mask as its alpha. JPEG 2000, JBIG2 and CCITT images are reported.
 */
import { encodePng } from '../drawingml/raster-encode';
import { jpegComponents } from '../drawingml/raster-decode';
import type { PdfDocument } from './document';
import { isName, isNumber, PdfName, PdfStream, PdfString, type PdfValue } from './objects';
import type { PdfPageImage } from './page-content';

const MAX_PIXELS = 40_000_000;

interface ColorSpace {
	components: number;
	toRgb(samples: number[], out: Uint8Array, at: number): void;
}

const gray: ColorSpace = {
	components: 1,
	toRgb: (s, out, at) => out.set([s[0]!, s[0]!, s[0]!], at),
};
const rgb: ColorSpace = {
	components: 3,
	toRgb: (s, out, at) => out.set([s[0]!, s[1]!, s[2]!], at),
};
const cmyk: ColorSpace = {
	components: 4,
	toRgb: (s, out, at) => {
		const k = 255 - s[3]!;
		out.set([((255 - s[0]!) * k) / 255, ((255 - s[1]!) * k) / 255, ((255 - s[2]!) * k) / 255], at);
	},
};

async function colorSpace(
	document: PdfDocument,
	value: PdfValue | undefined,
	depth = 0,
): Promise<ColorSpace | undefined> {
	const resolved = await document.resolve(value);
	if (depth > 4) return undefined;
	if (isName(resolved)) {
		if (['DeviceGray', 'G', 'CalGray'].includes(resolved.name)) return gray;
		if (['DeviceRGB', 'RGB', 'CalRGB'].includes(resolved.name)) return rgb;
		if (['DeviceCMYK', 'CMYK'].includes(resolved.name)) return cmyk;
		return undefined;
	}
	if (!Array.isArray(resolved) || !isName(resolved[0])) return undefined;
	const family = resolved[0].name;
	if (family === 'ICCBased') {
		const stream = await document.resolve(resolved[1]);
		const n = stream instanceof PdfStream ? stream.dict.get('N') : undefined;
		const alternate = stream instanceof PdfStream ? stream.dict.get('Alternate') : undefined;
		if (alternate) {
			const space = await colorSpace(document, alternate, depth + 1);
			if (space) return space;
		}
		return n === 1 ? gray : n === 4 ? cmyk : rgb;
	}
	if (family === 'CalRGB' || family === 'Lab') return rgb;
	if (family === 'CalGray') return gray;
	if (family === 'Indexed' || family === 'I') {
		const base = await colorSpace(document, resolved[1], depth + 1);
		const hival = await document.resolve(resolved[2]);
		const table = await document.resolve(resolved[3]);
		const lookup =
			table instanceof PdfString
				? table.bytes
				: table instanceof PdfStream
					? await document.streamBytes(table)
					: undefined;
		if (!base || !isNumber(hival) || !lookup) return undefined;
		return {
			components: 1,
			toRgb: (s, out, at) => {
				const index = Math.min(s[0]!, hival) * base.components;
				base.toRgb(Array.from(lookup.subarray(index, index + base.components)), out, at);
			},
		};
	}
	return undefined;
}

/** Samples of a decoded image, `bpc` bits each, scaled to 0..255 (palette indexes stay raw). */
function sampleReader(
	data: Uint8Array,
	width: number,
	components: number,
	bpc: number,
	indexed: boolean,
) {
	const stride = Math.ceil((width * components * bpc) / 8);
	return (x: number, y: number, component: number): number => {
		const index = x * components + component;
		if (bpc === 8) return data[y * stride + index] ?? 0;
		if (bpc === 16) return data[y * stride + index * 2] ?? 0;
		const bit = index * bpc;
		const byte = data[y * stride + (bit >> 3)] ?? 0;
		const value = (byte >> (8 - bpc - (bit & 7))) & ((1 << bpc) - 1);
		return indexed ? value : Math.round((value * 255) / ((1 << bpc) - 1));
	};
}

async function samples(document: PdfDocument, stream: PdfStream) {
	const { bytes, imageFilter } = await document.streamData(stream);
	return imageFilter ? undefined : bytes;
}

/** The picture an image XObject holds. */
export async function readImage(
	document: PdfDocument,
	stream: PdfStream,
	box: { x: number; y: number; width: number; height: number },
): Promise<PdfPageImage> {
	const dict = stream.dict;
	const width = (await document.resolve(dict.get('Width'))) as number;
	const height = (await document.resolve(dict.get('Height'))) as number;
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
	const decoded = await document.streamData(stream);
	const filter = decoded.imageFilter?.name;
	if (filter === 'DCTDecode') {
		const components = jpegComponents(decoded.bytes);
		if (components === 1 || components === 3)
			return { ...base, format: 'jpeg', data: decoded.bytes };
		return unsupported('it is a CMYK JPEG');
	}
	if (filter)
		return unsupported(`its ${filter.replace(/Decode$/, '')} compression is not supported`);
	const imageMask = (await document.resolve(dict.get('ImageMask'))) === true;
	const bpcValue = await document.resolve(dict.get('BitsPerComponent'));
	const bpc = imageMask ? 1 : isNumber(bpcValue) ? bpcValue : 8;
	const space = imageMask
		? gray
		: await colorSpace(document, dict.get('ColorSpace') ?? dict.get('CS'));
	if (!space) return unsupported('its colour space is not supported');
	const csValue = await document.resolve(dict.get('ColorSpace'));
	const indexed =
		Array.isArray(csValue) &&
		isName(csValue[0]) &&
		['Indexed', 'I'].includes((csValue[0] as PdfName).name);
	const read = sampleReader(decoded.bytes, width, space.components, bpc, indexed);
	const out = new Uint8Array(width * height * 3);
	const pixel: number[] = [];
	for (let y = 0; y < height; y++)
		for (let x = 0; x < width; x++) {
			for (let c = 0; c < space.components; c++) pixel[c] = read(x, y, c);
			space.toRgb(pixel, out, (y * width + x) * 3);
		}
	let alpha: Uint8Array | undefined;
	const smask = await document.resolve(dict.get('SMask'));
	if (smask instanceof PdfStream) {
		const maskWidth = await document.resolve(smask.dict.get('Width'));
		const maskHeight = await document.resolve(smask.dict.get('Height'));
		const maskBytes = await samples(document, smask);
		if (maskBytes && maskWidth === width && maskHeight === height) {
			const maskBpc = await document.resolve(smask.dict.get('BitsPerComponent'));
			const readMask = sampleReader(maskBytes, width, 1, isNumber(maskBpc) ? maskBpc : 8, false);
			alpha = new Uint8Array(width * height);
			for (let y = 0; y < height; y++)
				for (let x = 0; x < width; x++) alpha[y * width + x] = readMask(x, y, 0);
		}
	}
	if (imageMask) {
		// A stencil mask paints the fill colour where the sample is 0 (black ink on clear).
		const decode = await document.resolve(dict.get('Decode'));
		const paint = Array.isArray(decode) && decode[0] === 1 ? 255 : 0;
		alpha = new Uint8Array(width * height);
		for (let index = 0; index < alpha.length; index++) {
			alpha[index] = out[index * 3] === paint ? 255 : 0;
			out.fill(0, index * 3, index * 3 + 3);
		}
	}
	return { ...base, format: 'png', data: encodePng(width, height, out, alpha) };
}
