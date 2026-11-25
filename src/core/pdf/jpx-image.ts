/**
 * JPXDecode images (ISO 32000-1 section 7.4.9 and 8.9.5): the decoded JPEG 2000 components as
 * 8-bit samples at the image dictionary's size, in the colour space the PDF names (which then
 * overrides the JPEG 2000 header) or else the one the JP2 header gives, with the JP2 palette
 * applied (unless the PDF space is Indexed) and sub-sampled components spread over the grid.
 * With /SMaskInData the opacity channel becomes the alpha (premultiplied colour is undone).
 * /Decode and /BitsPerComponent are ignored, as the standard says, except for image masks.
 */
import type { PdfDocument } from './document';
import { applyIccProfile } from './icc-profile';
import { cmyk, colorSpace, gray, rgb, type ColorSpace } from './image-color';
import { decodeJpx, JpxError, type JpxComponent, type JpxImage } from './jpx/index';
import { isNumber, type PdfDict, type PdfValue } from './objects';

export interface JpxSamples {
	/** 8-bit samples, `space.components` per pixel, row by row. */
	bytes: Uint8Array;
	space: ColorSpace;
	/** The embedded opacity (with /SMaskInData), 0..255 per pixel. */
	alpha?: Uint8Array;
}

interface Channel {
	component: JpxComponent;
	/** Palette column values, indexed by the component's sample. */
	lookup?: Int32Array;
	depth: number;
	signed: boolean;
}

const ENUMERATED: Record<number, ColorSpace> = { 16: rgb, 20: rgb, 21: rgb, 17: gray, 12: cmyk };
const ICC: Record<string, ColorSpace> = { 'RGB ': rgb, GRAY: gray, CMYK: cmyk };
const YCC = new Set([18, 24, 1, 3, 4]);

/** The channels of the image: through the palette and component mapping, when there are any. */
function channels(image: JpxImage, usePalette: boolean): Channel[] {
	const { palette, mapping } = image.header;
	const direct = (component: JpxComponent): Channel => ({
		component,
		depth: component.depth,
		signed: component.signed,
	});
	if (!palette || !mapping || !usePalette) return image.components.map(direct);
	return mapping.map((map) => {
		const component = image.components[map.component];
		const column = palette.columns[map.column];
		if (!component) throw new JpxError('its JPEG 2000 component mapping is invalid');
		if (!map.palette || !column) return direct(component);
		return { component, lookup: column.values, depth: column.depth, signed: column.signed };
	});
}

/** Splits colour and opacity channels (by the channel definitions, or by count). */
function roles(
	list: Channel[],
	image: JpxImage,
	wanted: number | undefined,
	smaskInData: number,
): { colour: Channel[]; opacity?: Channel; premultiplied: boolean } {
	const defs = image.header.channels;
	if (defs?.length) {
		const colour: Channel[] = [];
		let opacity: Channel | undefined;
		let premultiplied = false;
		for (const def of defs) {
			const channel = list[def.channel];
			if (!channel) continue;
			if (def.type === 1 || def.type === 2) {
				if (!opacity && (def.association === 0 || def.association === 65535)) {
					opacity = channel;
					premultiplied = def.type === 2;
				}
			} else if (def.type === 0 && def.association >= 1 && def.association <= list.length)
				colour[def.association - 1] = channel;
		}
		const ordered = colour.filter(Boolean);
		if (ordered.length) return { colour: ordered, ...(opacity && { opacity }), premultiplied };
	}
	const expected = wanted ?? (list.length >= 4 ? (list.length === 4 && !smaskInData ? 4 : 3) : 1);
	if (smaskInData && list.length === expected + 1)
		return { colour: list.slice(0, expected), opacity: list[expected]!, premultiplied: false };
	return { colour: list, premultiplied: false };
}

/** The JPEG 2000 header's colour space for `count` colour channels. */
function headerSpace(image: JpxImage, count: number): { space: ColorSpace; ycc: boolean } {
	const { enumerated, icc } = image.header;
	const ycc = enumerated !== undefined && YCC.has(enumerated) && count >= 3;
	const named = ycc
		? rgb
		: enumerated !== undefined
			? ENUMERATED[enumerated]
			: icc
				? ICC[icc]
				: undefined;
	if (named && named.components <= count) return { space: named, ycc };
	return { space: count >= 4 ? cmyk : count >= 3 ? rgb : gray, ycc: false };
}

/** Per output column (or row), the sample index of a component, nearest on the reference grid. */
function positions(size: number, target: number, origin: number, step: number, count: number) {
	const first = Math.ceil(origin / step);
	return Int32Array.from({ length: target }, (_, index) => {
		const grid = origin + Math.floor(((index + 0.5) * size) / target);
		return Math.min(count - 1, Math.max(0, Math.floor(grid / step) - first));
	});
}

/** Decodes a JPXDecode image's data for an image dictionary of `width` x `height`. */
export async function jpxSamples(
	document: PdfDocument,
	source: { dict: PdfDict; named?: (name: string) => Promise<PdfValue | undefined> },
	data: Uint8Array,
	width: number,
	height: number,
	asMask = false,
): Promise<JpxSamples | string> {
	const get = (key: string) => document.resolve(source.dict.get(key));
	let image: JpxImage;
	try {
		image = decodeJpx(data, { maxSamples: Math.max(1, width * height) * 5 });
	} catch (error) {
		return error instanceof JpxError ? error.message : 'its JPEG 2000 data could not be read';
	}
	const imageMask = (await get('ImageMask')) === true;
	const named = asMask || imageMask ? undefined : await get('ColorSpace');
	const pdfSpace =
		named === undefined ? undefined : await colorSpace(document, named, source.named);
	if (named !== undefined && !pdfSpace) return 'its colour space is not supported';
	const indexed = pdfSpace?.hival !== undefined;
	const smaskValue = await get('SMaskInData');
	const smaskInData = isNumber(smaskValue) && !asMask ? smaskValue : 0;
	const list = channels(image, !indexed);
	const wanted = asMask || imageMask ? 1 : pdfSpace?.components;
	const { colour, opacity, premultiplied } = roles(list, image, wanted, smaskInData);
	let space: ColorSpace;
	let ycc = false;
	if (pdfSpace) space = pdfSpace;
	else if (asMask || imageMask) space = gray;
	else ({ space, ycc } = headerSpace(image, colour.length));
	if (colour.length < space.components)
		return 'its JPEG 2000 channels do not match its colour space';
	const used = [...colour.slice(0, space.components), ...(smaskInData && opacity ? [opacity] : [])];
	const n = used.length;
	const out = new Uint8Array(width * height * n);
	used.forEach((channel, c) => {
		const { component } = channel;
		const xs = positions(image.width, width, image.x0, component.dx, component.width);
		const ys = positions(image.height, height, image.y0, component.dy, component.height);
		const max = 2 ** channel.depth - 1;
		const offset = channel.signed ? 2 ** (channel.depth - 1) : 0;
		const raw = indexed && c < space.components;
		for (let y = 0; y < height; y++) {
			const row = ys[y]! * component.width;
			for (let x = 0; x < width; x++) {
				let value = component.data[row + xs[x]!]!;
				if (channel.lookup)
					value = channel.lookup[Math.min(channel.lookup.length - 1, Math.max(0, value))]!;
				out[(y * width + x) * n + c] = raw
					? Math.min(255, Math.max(0, value))
					: Math.round(((value + offset) * 255) / max);
			}
		}
	});
	if (ycc) toRgb(out, n);
	else if (!pdfSpace && !asMask && !imageMask && image.header.iccProfile)
		applyIccProfile(out, n, image.header.iccProfile);
	const pixels = width * height;
	const components = space.components;
	if (n === components) return { bytes: out, space };
	const bytes = new Uint8Array(pixels * components);
	const alpha = new Uint8Array(pixels);
	const unpremultiply = premultiplied || smaskInData === 2;
	for (let i = 0; i < pixels; i++) {
		const a = out[i * n + components]!;
		alpha[i] = a;
		for (let c = 0; c < components; c++) {
			const value = out[i * n + c]!;
			bytes[i * components + c] =
				unpremultiply && a ? Math.min(255, Math.round((value * 255) / a)) : value;
		}
	}
	return { bytes, space, alpha };
}

/** sYCC (and the other YCbCr spaces) to RGB, in place on the first three of `n` samples. */
function toRgb(out: Uint8Array, n: number): void {
	for (let at = 0; at + 2 < out.length; at += n) {
		const y = out[at]!;
		const cb = out[at + 1]! - 128;
		const cr = out[at + 2]! - 128;
		out[at] = clamp(y + 1.402 * cr);
		out[at + 1] = clamp(y - 0.344136 * cb - 0.714136 * cr);
		out[at + 2] = clamp(y + 1.772 * cb);
	}
}

const clamp = (value: number) => Math.min(255, Math.max(0, Math.round(value)));
