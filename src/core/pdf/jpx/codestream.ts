/**
 * Splits a JPEG 2000 codestream (ITU-T T.800 Annex A) into its main header, and per tile the
 * markers of its tile-part headers, the packet data of its tile-parts in order and the packed
 * packet headers (PPM in the main header, PPT in the tile-part headers) when there are any.
 */
import { JpxError, type JpxLimits } from './errors';
import { emptyMarkers, readMarker, readSiz, type Markers, type Siz } from './markers';

export interface TileData {
	index: number;
	markers: Markers;
	/** Packet data of each tile-part, in codestream order. */
	parts: Uint8Array[];
	/** Packed packet headers (PPM or PPT), in order, when the codestream has them. */
	headers?: Uint8Array[];
}

export interface Codestream {
	siz: Siz;
	main: Markers;
	tiles: TileData[];
}

const SOC = 0xff4f;
const SIZ = 0xff51;
const SOT = 0xff90;
const SOD = 0xff93;
const EOC = 0xffd9;
const PPM = 0xff60;
const PPT = 0xff61;

const u16 = (data: Uint8Array, at: number) => (data[at]! << 8) | data[at + 1]!;
const u32 = (data: Uint8Array, at: number) => u16(data, at) * 65536 + u16(data, at + 2);

/** Reads the marker segments from `at` up to SOT or SOD; returns the position of that marker. */
function segments(
	data: Uint8Array,
	at: number,
	each: (code: number, body: Uint8Array) => void,
): number {
	for (;;) {
		if (at + 4 > data.length) throw new JpxError('its JPEG 2000 header is truncated');
		const code = u16(data, at);
		if (code === SOT || code === SOD) return at;
		if (code >> 8 !== 0xff) throw new JpxError('its JPEG 2000 header is damaged');
		const length = u16(data, at + 2);
		if (length < 2 || at + 2 + length > data.length)
			throw new JpxError('its JPEG 2000 header is truncated');
		each(code, data.subarray(at + 4, at + 2 + length));
		at += 2 + length;
	}
}

/** Packed headers in Z order, with their index byte removed. */
const packed = (list: { z: number; data: Uint8Array }[]) =>
	list.sort((a, b) => a.z - b.z).map((item) => item.data);

const concat = (parts: Uint8Array[]): Uint8Array => {
	if (parts.length === 1) return parts[0]!;
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
};

export function readCodestream(data: Uint8Array, limits: JpxLimits): Codestream {
	if (data.length < 4 || u16(data, 0) !== SOC || u16(data, 2) !== SIZ)
		throw new JpxError('its JPEG 2000 codestream is invalid');
	const sizLength = u16(data, 4);
	const siz = readSiz(data.subarray(6, 4 + sizLength), limits);
	const count = siz.components.length;
	const main = emptyMarkers();
	const ppm: { z: number; data: Uint8Array }[] = [];
	let at = segments(data, 4 + sizLength, (code, body) => {
		if (code === PPM && body.length) ppm.push({ z: body[0]!, data: body.subarray(1) });
		else readMarker(code, body, main, count, limits);
	});
	if (!main.tileStyle || !main.style || !main.quantization)
		throw new JpxError('its JPEG 2000 header lacks the coding style or quantization');
	const tiles = new Map<number, TileData>();
	const ppt = new Map<TileData, { z: number; data: Uint8Array }[]>();
	const partOrder: TileData[] = [];
	const total = siz.tilesWide * siz.tilesHigh;
	while (at + 12 <= data.length && u16(data, at) === SOT) {
		const index = u16(data, at + 4);
		const length = u32(data, at + 6);
		const partIndex = data[at + 10]!;
		if (index >= total) throw new JpxError('its JPEG 2000 tile index is invalid');
		let tile = tiles.get(index);
		if (!tile) {
			tile = { index, markers: emptyMarkers(), parts: [] };
			tiles.set(index, tile);
		}
		const target = tile;
		const sod = segments(data, at + 2 + u16(data, at + 2), (code, body) => {
			if (code === PPT && body.length)
				ppt.set(target, [...(ppt.get(target) ?? []), { z: body[0]!, data: body.subarray(1) }]);
			// Only the first tile-part may change the coding style or quantization.
			else if (partIndex === 0 || code === 0xff5f)
				readMarker(code, body, target.markers, count, limits);
		});
		if (u16(data, sod) !== SOD) break;
		let end = length ? at + length : data.length;
		if (!length && u16(data, data.length - 2) === EOC) end -= 2;
		end = Math.min(end, data.length);
		tile.parts.push(data.subarray(sod + 2, Math.max(sod + 2, end)));
		partOrder.push(tile);
		at = end;
	}
	if (!tiles.size) throw new JpxError('its JPEG 2000 codestream has no tiles');
	if (ppm.length) {
		// Nppm and Ippm records, one per tile-part in codestream order.
		const stream = concat(packed(ppm));
		const headers = new Map<TileData, Uint8Array[]>();
		let position = 0;
		for (const tile of partOrder) {
			if (position + 4 > stream.length) break;
			const size = u32(stream, position);
			position += 4;
			const list = headers.get(tile) ?? [];
			list.push(stream.subarray(position, Math.min(stream.length, position + size)));
			headers.set(tile, list);
			position += size;
		}
		for (const [tile, list] of headers) tile.headers = list;
	}
	for (const [tile, list] of ppt) tile.headers = packed(list);
	return {
		siz,
		main,
		tiles: [...tiles.values()],
	};
}
