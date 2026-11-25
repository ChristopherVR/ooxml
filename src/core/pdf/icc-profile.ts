/**
 * Lightweight ICC v2/v4 Matrix/TRC profile converter to sRGB (ISO 15076-1).
 * Handles RGB Matrix/TRC (`rXYZ`, `gXYZ`, `bXYZ`, `rTRC`, `gTRC`, `bTRC`) and Monochrome (`kTRC`)
 * profiles embedded in JPEG 2000 `colr` boxes, converting PCS D50 XYZ to D65 sRGB. DOM-free.
 */

const u16 = (d: Uint8Array, at: number) => (d[at]! << 8) | d[at + 1]!;
const u32 = (d: Uint8Array, at: number) => u16(d, at) * 65536 + u16(d, at + 2);
const s15Fixed16 = (d: Uint8Array, at: number) => {
	const raw = (d[at]! << 24) | (d[at + 1]! << 16) | (d[at + 2]! << 8) | d[at + 3]!;
	return raw / 65536;
};
const sig = (d: Uint8Array, at: number) =>
	String.fromCharCode(d[at]!, d[at + 1]!, d[at + 2]!, d[at + 3]!);

/** Evaluates an ICC `curv` or `para` tag into a 256-entry linearisation LUT (0..1). */
function readTrcLut(profile: Uint8Array, offset: number, size: number): Float32Array | undefined {
	if (offset + 12 > profile.length || size < 12) return undefined;
	const tagType = sig(profile, offset);
	const lut = new Float32Array(256);
	if (tagType === 'curv') {
		const count = u32(profile, offset + 8);
		if (count === 0) {
			for (let i = 0; i < 256; i++) lut[i] = i / 255;
			return lut;
		}
		if (count === 1) {
			if (offset + 14 > profile.length) return undefined;
			const gamma = u16(profile, offset + 12) / 256;
			for (let i = 0; i < 256; i++) lut[i] = Math.pow(i / 255, gamma);
			return lut;
		}
		if (offset + 12 + count * 2 > profile.length) return undefined;
		for (let i = 0; i < 256; i++) {
			const pos = (i / 255) * (count - 1);
			const lo = Math.floor(pos);
			const hi = Math.min(count - 1, lo + 1);
			const frac = pos - lo;
			const v0 = u16(profile, offset + 12 + lo * 2) / 65535;
			const v1 = u16(profile, offset + 12 + hi * 2) / 65535;
			lut[i] = v0 + frac * (v1 - v0);
		}
		return lut;
	}
	if (tagType === 'para') {
		if (offset + 16 > profile.length) return undefined;
		const fnType = u16(profile, offset + 8);
		const paramCounts = [1, 3, 4, 5, 7];
		const count = paramCounts[fnType];
		if (!count || offset + 12 + count * 4 > profile.length) return undefined;
		const p = Array.from({ length: count }, (_, idx) => s15Fixed16(profile, offset + 12 + idx * 4));
		const [g = 1, a = 1, b = 0, c = 0, d = 0, e = 0, f = 0] = p;
		for (let i = 0; i < 256; i++) {
			const x = i / 255;
			let y: number;
			if (fnType === 0) y = Math.pow(x, g);
			else if (fnType === 1) y = x >= -b / a ? Math.pow(a * x + b, g) : 0;
			else if (fnType === 2) y = x >= -b / a ? Math.pow(a * x + b, g) + c : c;
			else if (fnType === 3) y = x >= d ? Math.pow(a * x + b, g) : c * x;
			else y = x >= d ? Math.pow(a * x + b, g) + e : c * x + f;
			lut[i] = Math.max(0, Math.min(1, y));
		}
		return lut;
	}
	return undefined;
}

function readXyz(
	profile: Uint8Array,
	offset: number,
	size: number,
): [number, number, number] | undefined {
	if (size < 20 || offset + 20 > profile.length || sig(profile, offset) !== 'XYZ ')
		return undefined;
	return [
		s15Fixed16(profile, offset + 8),
		s15Fixed16(profile, offset + 12),
		s15Fixed16(profile, offset + 16),
	];
}

/** Linear sRGB (0..1) to 8-bit sRGB (0..255). */
function linearToSrgbByte(value: number): number {
	const v = Math.max(0, Math.min(1, value));
	const encoded = v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
	return Math.max(0, Math.min(255, Math.round(encoded * 255)));
}

/**
 * Combined PCS D50 XYZ -> Bradford D65 -> linear sRGB 3x3 matrix.
 * (Standard ICC D50 to sRGB D65 matrix.)
 */
const XYZ_D50_TO_SRGB = [
	[3.1338561, -1.6168667, -0.4906146],
	[-0.9787684, 1.9161415, 0.033454],
	[0.0719453, -0.2289914, 1.4052427],
] as const;

/**
 * Applies an embedded ICC Matrix/TRC RGB or Monochrome TRC profile in place to `samples`
 * (8-bit packed, `components` per pixel). Returns whether the profile was applied.
 */
export function applyIccProfile(
	samples: Uint8Array,
	components: number,
	profile: Uint8Array,
): boolean {
	if (profile.length < 132 || sig(profile, 36) !== 'acsp') return false;
	const space = sig(profile, 16);
	const tagCount = Math.min(256, u32(profile, 128));
	if (132 + tagCount * 12 > profile.length) return false;
	const tags = new Map<string, { offset: number; size: number }>();
	for (let i = 0; i < tagCount; i++) {
		const at = 132 + i * 12;
		tags.set(sig(profile, at), { offset: u32(profile, at + 4), size: u32(profile, at + 8) });
	}
	if (space === 'GRAY' && components === 1) {
		const kTag = tags.get('kTRC');
		const lut = kTag && readTrcLut(profile, kTag.offset, kTag.size);
		if (!lut) return false;
		const map = new Uint8Array(256);
		for (let i = 0; i < 256; i++) map[i] = linearToSrgbByte(lut[i]!);
		for (let i = 0; i < samples.length; i++) samples[i] = map[samples[i]!]!;
		return true;
	}
	if (space === 'RGB ' && components >= 3) {
		const rXyzTag = tags.get('rXYZ');
		const gXyzTag = tags.get('gXYZ');
		const bXyzTag = tags.get('bXYZ');
		const rTrcTag = tags.get('rTRC');
		const gTrcTag = tags.get('gTRC');
		const bTrcTag = tags.get('bTRC');
		if (!rXyzTag || !gXyzTag || !bXyzTag || !rTrcTag || !gTrcTag || !bTrcTag) return false;
		const rXyz = readXyz(profile, rXyzTag.offset, rXyzTag.size);
		const gXyz = readXyz(profile, gXyzTag.offset, gXyzTag.size);
		const bXyz = readXyz(profile, bXyzTag.offset, bXyzTag.size);
		const rLut = readTrcLut(profile, rTrcTag.offset, rTrcTag.size);
		const gLut = readTrcLut(profile, gTrcTag.offset, gTrcTag.size);
		const bLut = readTrcLut(profile, bTrcTag.offset, bTrcTag.size);
		if (!rXyz || !gXyz || !bXyz || !rLut || !gLut || !bLut) return false;
		// Compose M = XYZ_D50_TO_SRGB * [rXyz gXyz bXyz] (columns).
		const m = [0, 1, 2].map((row) => {
			const [a, b, c] = XYZ_D50_TO_SRGB[row]!;
			return [
				a * rXyz[0] + b * rXyz[1] + c * rXyz[2],
				a * gXyz[0] + b * gXyz[1] + c * gXyz[2],
				a * bXyz[0] + b * bXyz[1] + c * bXyz[2],
			] as const;
		});
		const [m0, m1, m2] = m as [
			readonly [number, number, number],
			readonly [number, number, number],
			readonly [number, number, number],
		];
		for (let i = 0; i + 2 < samples.length; i += components) {
			const lr = rLut[samples[i]!]!;
			const lg = gLut[samples[i + 1]!]!;
			const lb = bLut[samples[i + 2]!]!;
			samples[i] = linearToSrgbByte(m0[0] * lr + m0[1] * lg + m0[2] * lb);
			samples[i + 1] = linearToSrgbByte(m1[0] * lr + m1[1] * lg + m1[2] * lb);
			samples[i + 2] = linearToSrgbByte(m2[0] * lr + m2[1] * lg + m2[2] * lb);
		}
		return true;
	}
	return false;
}
