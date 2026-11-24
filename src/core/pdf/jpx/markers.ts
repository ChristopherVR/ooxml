/**
 * The marker segments of a JPEG 2000 codestream header (ITU-T T.800 A.5 and A.6): SIZ, COD, COC,
 * QCD, QCC, RGN and POC, read into plain records. Other markers (COM, TLM, PLM, PLT, CRG) carry
 * nothing the decoder needs and are skipped; PPM and PPT are collected by the codestream reader.
 */
import { JpxError, type JpxLimits } from './errors';

export interface JpxComponentInfo {
	depth: number;
	signed: boolean;
	dx: number;
	dy: number;
}

export interface Siz {
	/** Reference grid: image area [x0, x1) x [y0, y1). */
	x0: number;
	y0: number;
	x1: number;
	y1: number;
	tileWidth: number;
	tileHeight: number;
	tileX0: number;
	tileY0: number;
	tilesWide: number;
	tilesHigh: number;
	components: JpxComponentInfo[];
}

/** The coding style of one component (SPcod / SPcoc). */
export interface ComponentStyle {
	levels: number;
	/** Code-block width and height exponents. */
	xcb: number;
	ycb: number;
	/** Code-block style flags (bypass 1, reset 2, termall 4, vcausal 8, predictable 16, segsym 32). */
	style: number;
	reversible: boolean;
	/** Precinct size exponents per resolution level, 0..levels. */
	ppx: number[];
	ppy: number[];
}

/** The coding style that applies to the whole tile (Scod / SGcod). */
export interface TileStyle {
	sop: boolean;
	eph: boolean;
	order: number;
	layers: number;
	mct: number;
}

export interface Quantization {
	/** 0 none (reversible), 1 scalar derived, 2 scalar expounded. */
	style: number;
	guard: number;
	exponents: number[];
	mantissas: number[];
}

export interface ProgressionChange {
	resolutionStart: number;
	componentStart: number;
	layerEnd: number;
	resolutionEnd: number;
	componentEnd: number;
	order: number;
}

/** The markers of the main header or of a tile's headers. */
export interface Markers {
	tileStyle?: TileStyle;
	style?: ComponentStyle;
	componentStyles: Map<number, ComponentStyle>;
	quantization?: Quantization;
	componentQuantization: Map<number, Quantization>;
	roiShift: Map<number, number>;
	progression: ProgressionChange[];
}

export const emptyMarkers = (): Markers => ({
	componentStyles: new Map(),
	componentQuantization: new Map(),
	roiShift: new Map(),
	progression: [],
});

const u16 = (data: Uint8Array, at: number) => {
	if (at + 2 > data.length) throw new JpxError('its JPEG 2000 header is truncated');
	return (data[at]! << 8) | data[at + 1]!;
};
const u32 = (data: Uint8Array, at: number) => u16(data, at) * 65536 + u16(data, at + 2);
const u8 = (data: Uint8Array, at: number) => {
	if (at >= data.length) throw new JpxError('its JPEG 2000 header is truncated');
	return data[at]!;
};

export function readSiz(data: Uint8Array, limits: JpxLimits): Siz {
	const count = u16(data, 34);
	const siz = {
		x1: u32(data, 2),
		y1: u32(data, 6),
		x0: u32(data, 10),
		y0: u32(data, 14),
		tileWidth: u32(data, 18),
		tileHeight: u32(data, 22),
		tileX0: u32(data, 26),
		tileY0: u32(data, 30),
		components: Array.from({ length: count }, (_, index) => {
			const ssiz = u8(data, 36 + index * 3);
			return {
				depth: (ssiz & 0x7f) + 1,
				signed: (ssiz & 0x80) !== 0,
				dx: u8(data, 37 + index * 3),
				dy: u8(data, 38 + index * 3),
			};
		}),
	};
	const bad =
		!count ||
		count > 16384 ||
		siz.x1 <= siz.x0 ||
		siz.y1 <= siz.y0 ||
		!siz.tileWidth ||
		!siz.tileHeight ||
		siz.tileX0 > siz.x0 ||
		siz.tileY0 > siz.y0 ||
		siz.tileX0 + siz.tileWidth <= siz.x0 ||
		siz.tileY0 + siz.tileHeight <= siz.y0 ||
		siz.components.some((c) => !c.dx || !c.dy || c.depth > 38);
	if (bad) throw new JpxError('its JPEG 2000 image size is invalid');
	const tilesWide = Math.ceil((siz.x1 - siz.tileX0) / siz.tileWidth);
	const tilesHigh = Math.ceil((siz.y1 - siz.tileY0) / siz.tileHeight);
	if (tilesWide * tilesHigh > limits.maxTiles)
		throw new JpxError('its JPEG 2000 image has too many tiles');
	const samples = siz.components.reduce(
		(sum, c) => sum + Math.ceil((siz.x1 - siz.x0) / c.dx) * Math.ceil((siz.y1 - siz.y0) / c.dy),
		0,
	);
	if (samples > limits.maxSamples) throw new JpxError('it is larger than the size limit');
	return { ...siz, tilesWide, tilesHigh };
}

/** SPcod / SPcoc at `at`; `precincts` when the precinct sizes follow. */
function componentStyle(data: Uint8Array, at: number, precincts: boolean): ComponentStyle {
	const levels = u8(data, at);
	const xcb = (u8(data, at + 1) & 0x0f) + 2;
	const ycb = (u8(data, at + 2) & 0x0f) + 2;
	const style = u8(data, at + 3);
	const transform = u8(data, at + 4);
	if (levels > 32 || xcb > 10 || ycb > 10 || xcb + ycb > 12 || transform > 1)
		throw new JpxError('its JPEG 2000 coding style is invalid');
	if (style & 0x40)
		throw new JpxError('its High Throughput JPEG 2000 coding (Part 15) is not supported');
	const ppx: number[] = [];
	const ppy: number[] = [];
	for (let r = 0; r <= levels; r++) {
		const size = precincts ? u8(data, at + 5 + r) : 0xff;
		ppx.push(size & 0x0f);
		ppy.push(size >> 4);
	}
	return { levels, xcb, ycb, style, reversible: transform === 1, ppx, ppy };
}

function quantization(data: Uint8Array, at: number, end: number): Quantization {
	const sqcd = u8(data, at);
	const style = sqcd & 0x1f;
	if (style > 2) throw new JpxError('its JPEG 2000 quantization is invalid');
	const exponents: number[] = [];
	const mantissas: number[] = [];
	for (let position = at + 1; position < end; position += style ? 2 : 1) {
		if (style === 0) {
			exponents.push(u8(data, position) >> 3);
			mantissas.push(0);
		} else {
			const value = u16(data, position);
			exponents.push(value >> 11);
			mantissas.push(value & 0x7ff);
			if (style === 1) break;
		}
	}
	if (!exponents.length) throw new JpxError('its JPEG 2000 quantization is invalid');
	return { style, guard: sqcd >> 5, exponents, mantissas };
}

/**
 * Reads one marker segment (`code`, body `data` without the length) into `markers`. Returns
 * false for markers it does not keep.
 */
export function readMarker(
	code: number,
	data: Uint8Array,
	markers: Markers,
	components: number,
	limits: JpxLimits,
): boolean {
	const wide = components > 256;
	const component = (at: number) => (wide ? u16(data, at) : u8(data, at));
	const step = wide ? 2 : 1;
	const checkComponent = (index: number) => {
		if (index >= components) throw new JpxError('its JPEG 2000 header names a missing component');
		return index;
	};
	switch (code) {
		case 0xff52: {
			const scod = u8(data, 0);
			const layers = u16(data, 2);
			const order = u8(data, 1);
			if (!layers || layers > limits.maxLayers || order > 4)
				throw new JpxError('its JPEG 2000 coding style is invalid');
			markers.tileStyle = {
				sop: (scod & 2) !== 0,
				eph: (scod & 4) !== 0,
				order,
				layers,
				mct: u8(data, 4),
			};
			markers.style = componentStyle(data, 5, (scod & 1) !== 0);
			return true;
		}
		case 0xff53: {
			const index = checkComponent(component(0));
			markers.componentStyles.set(
				index,
				componentStyle(data, step + 1, (u8(data, step) & 1) !== 0),
			);
			return true;
		}
		case 0xff5c:
			markers.quantization = quantization(data, 0, data.length);
			return true;
		case 0xff5d: {
			const index = checkComponent(component(0));
			markers.componentQuantization.set(index, quantization(data, step, data.length));
			return true;
		}
		case 0xff5e: {
			const index = checkComponent(component(0));
			if (u8(data, step) !== 0 || u8(data, step + 1) > 37)
				throw new JpxError('its JPEG 2000 region of interest is invalid');
			markers.roiShift.set(index, u8(data, step + 1));
			return true;
		}
		case 0xff5f: {
			const size = 5 + 2 * step;
			for (let at = 0; at + size <= data.length; at += size) {
				const componentEnd = component(at + 4 + step);
				markers.progression.push({
					resolutionStart: u8(data, at),
					componentStart: component(at + 1),
					layerEnd: u16(data, at + 1 + step),
					resolutionEnd: u8(data, at + 3 + step),
					componentEnd: componentEnd === 0 && !wide ? 256 : componentEnd,
					order: u8(data, at + 4 + 2 * step),
				});
			}
			return true;
		}
		default:
			return false;
	}
}
