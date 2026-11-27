/**
 * A small ICC v2 display profile for sRGB (IEC 61966-2.1), built from the published colorimetry
 * rather than shipped or downloaded: the D50-adapted primaries (Bradford), the D65 white point
 * and the sRGB transfer curve sampled at 1024 points. PDF/A output intents embed it. DOM-free.
 */

const D50 = [0.9642, 1, 0.8249] as const;
const D65 = [0.9505, 1, 1.089] as const;
const RED = [0.4361, 0.2225, 0.0139] as const;
const GREEN = [0.3851, 0.7169, 0.0971] as const;
const BLUE = [0.1431, 0.0606, 0.7141] as const;
const CURVE_POINTS = 1024;

class Writer {
	readonly bytes: number[] = [];
	ascii(text: string): this {
		for (const character of text) this.bytes.push(character.charCodeAt(0) & 0x7f);
		return this;
	}
	u8(value: number): this {
		this.bytes.push(value & 0xff);
		return this;
	}
	u16(value: number): this {
		return this.u8(value >>> 8).u8(value);
	}
	u32(value: number): this {
		return this.u16(value >>> 16).u16(value & 0xffff);
	}
	s15f16(value: number): this {
		return this.u32(Math.round(value * 65536) >>> 0);
	}
	zeros(count: number): this {
		for (let index = 0; index < count; index++) this.bytes.push(0);
		return this;
	}
	pad(): this {
		while (this.bytes.length % 4) this.bytes.push(0);
		return this;
	}
}

const xyzTag = (xyz: readonly number[]) => {
	const tag = new Writer().ascii('XYZ ').zeros(4);
	for (const value of xyz) tag.s15f16(value);
	return tag.bytes;
};

const textTag = (text: string) => new Writer().ascii('text').zeros(4).ascii(text).u8(0).bytes;

/** `textDescriptionType`: the ASCII name, then empty Unicode and ScriptCode names. */
const descriptionTag = (text: string) =>
	new Writer()
		.ascii('desc')
		.zeros(4)
		.u32(text.length + 1)
		.ascii(text)
		.u8(0)
		.u32(0)
		.u32(0)
		.u16(0)
		.u8(0)
		.zeros(67).bytes;

const curveTag = () => {
	const tag = new Writer().ascii('curv').zeros(4).u32(CURVE_POINTS);
	for (let index = 0; index < CURVE_POINTS; index++) {
		const v = index / (CURVE_POINTS - 1);
		const linear = v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
		tag.u16(Math.round(linear * 65535));
	}
	return tag.bytes;
};

let cached: Uint8Array | undefined;

/** The profile's bytes (about 2.6 KB), built once. */
export function srgbIccProfile(): Uint8Array {
	if (cached) return cached;
	const curve = curveTag();
	const tags: [string, number[]][] = [
		['desc', descriptionTag('sRGB IEC61966-2.1')],
		['cprt', textTag('No copyright, use freely')],
		['wtpt', xyzTag(D65)],
		['rXYZ', xyzTag(RED)],
		['gXYZ', xyzTag(GREEN)],
		['bXYZ', xyzTag(BLUE)],
		['rTRC', curve],
		['gTRC', curve],
		['bTRC', curve],
	];
	const tableSize = 4 + tags.length * 12;
	const body = new Writer();
	const entries: { signature: string; offset: number; size: number }[] = [];
	const offsets = new Map<number[], number>();
	for (const [signature, data] of tags) {
		let offset = offsets.get(data);
		if (offset === undefined) {
			body.pad();
			offset = 128 + tableSize + body.bytes.length;
			offsets.set(data, offset);
			body.bytes.push(...data);
		}
		entries.push({ signature, offset, size: data.length });
	}
	body.pad();
	const size = 128 + tableSize + body.bytes.length;
	const header = new Writer()
		.u32(size)
		.zeros(4)
		.u32(0x02100000)
		.ascii('mntrRGB XYZ ')
		.u16(2026)
		.u16(1)
		.u16(1)
		.zeros(6)
		.ascii('acsp')
		.zeros(4 + 4 + 4 + 4 + 8 + 4);
	for (const value of D50) header.s15f16(value);
	header.zeros(128 - header.bytes.length);
	const table = new Writer().u32(tags.length);
	for (const entry of entries) table.ascii(entry.signature).u32(entry.offset).u32(entry.size);
	cached = Uint8Array.from([...header.bytes, ...table.bytes, ...body.bytes]);
	return cached;
}
