/**
 * Image colour spaces (ISO 32000-1 section 8.6) reduced to what a picture needs: the component
 * count and a conversion of 0..255 samples to RGB. DeviceN, Separation and Lab are approximated
 * through their alternate space or as RGB.
 */
import type { PdfDocument } from './document';
import { isName, isNumber, PdfStream, PdfString, type PdfValue } from './objects';

export interface ColorSpace {
	components: number;
	/** Converts samples (0..255, or palette indexes for an indexed space) to RGB at `at`. */
	toRgb(samples: number[], out: Uint8Array, at: number): void;
	/** The largest index of an indexed space. */
	hival?: number;
}

export const gray: ColorSpace = {
	components: 1,
	toRgb: (s, out, at) => out.set([s[0]!, s[0]!, s[0]!], at),
};
export const rgb: ColorSpace = {
	components: 3,
	toRgb: (s, out, at) => out.set([s[0]!, s[1]!, s[2]!], at),
};
export const cmyk: ColorSpace = {
	components: 4,
	toRgb: (s, out, at) => {
		const k = 255 - s[3]!;
		out.set([((255 - s[0]!) * k) / 255, ((255 - s[1]!) * k) / 255, ((255 - s[2]!) * k) / 255], at);
	},
};

const DEVICE: Record<string, ColorSpace> = {
	DeviceGray: gray,
	G: gray,
	CalGray: gray,
	DeviceRGB: rgb,
	RGB: rgb,
	CalRGB: rgb,
	DeviceCMYK: cmyk,
	CMYK: cmyk,
};

/** Resolves a colour space; `named` looks up other names (an inline image's resources). */
export async function colorSpace(
	document: PdfDocument,
	value: PdfValue | undefined,
	named?: (name: string) => Promise<PdfValue | undefined>,
	depth = 0,
): Promise<ColorSpace | undefined> {
	if (depth > 4) return undefined;
	const resolved = await document.resolve(value);
	if (isName(resolved)) {
		const device = DEVICE[resolved.name];
		if (device) return device;
		const lookup = named ? await named(resolved.name) : undefined;
		return lookup === undefined ? undefined : colorSpace(document, lookup, named, depth + 1);
	}
	if (!Array.isArray(resolved) || !isName(resolved[0])) return undefined;
	const family = resolved[0].name;
	if (family === 'ICCBased') {
		const stream = await document.resolve(resolved[1]);
		const n =
			stream instanceof PdfStream ? await document.resolve(stream.dict.get('N')) : undefined;
		const alternate = stream instanceof PdfStream ? stream.dict.get('Alternate') : undefined;
		if (alternate) {
			const space = await colorSpace(document, alternate, named, depth + 1);
			if (space) return space;
		}
		return n === 1 ? gray : n === 4 ? cmyk : rgb;
	}
	if (family === 'CalRGB' || family === 'Lab') return rgb;
	if (family === 'CalGray') return gray;
	if (family === 'Separation' || family === 'DeviceN') {
		// Without the tint transform (a function), a single colorant prints as a grey ramp.
		const names = await document.resolve(resolved[1]);
		const count = family === 'DeviceN' && Array.isArray(names) ? names.length : 1;
		if (count === 1)
			return { components: 1, toRgb: (s, out, at) => out.fill(255 - s[0]!, at, at + 3) };
		return count === 4 ? cmyk : count === 3 ? rgb : undefined;
	}
	if (family === 'Indexed' || family === 'I') {
		const base = await colorSpace(document, resolved[1], named, depth + 1);
		const hival = await document.resolve(resolved[2]);
		const table = await document.resolve(resolved[3]);
		const lookup =
			table instanceof PdfString
				? table.bytes
				: table instanceof PdfStream
					? await document.streamBytes(table)
					: undefined;
		if (!base || !isNumber(hival) || !lookup) return undefined;
		const pixel: number[] = [];
		return {
			components: 1,
			hival,
			toRgb: (s, out, at) => {
				const index = Math.min(s[0]!, hival) * base.components;
				for (let c = 0; c < base.components; c++) pixel[c] = lookup[index + c] ?? 0;
				base.toRgb(pixel, out, at);
			},
		};
	}
	return undefined;
}
