/**
 * A minimal JBIG2 encoder for tests (ITU-T T.88): the MQ encoder of Annex E.2 (its probability
 * table copied separately from the decoder's), the integer encoder of A.2 and segment builders.
 * The region encoders are in `jbig2-encode-regions.ts`.
 */
const QE = [
	[0x5601, 1, 1, 1],
	[0x3401, 2, 6, 0],
	[0x1801, 3, 9, 0],
	[0x0ac1, 4, 12, 0],
	[0x0521, 5, 29, 0],
	[0x0221, 38, 33, 0],
	[0x5601, 7, 6, 1],
	[0x5401, 8, 14, 0],
	[0x4801, 9, 14, 0],
	[0x3801, 10, 14, 0],
	[0x3001, 11, 17, 0],
	[0x2401, 12, 18, 0],
	[0x1c01, 13, 20, 0],
	[0x1601, 29, 21, 0],
	[0x5601, 15, 14, 1],
	[0x5401, 16, 14, 0],
	[0x5101, 17, 15, 0],
	[0x4801, 18, 16, 0],
	[0x3801, 19, 17, 0],
	[0x3401, 20, 18, 0],
	[0x3001, 21, 19, 0],
	[0x2801, 22, 19, 0],
	[0x2401, 23, 20, 0],
	[0x2201, 24, 21, 0],
	[0x1c01, 25, 22, 0],
	[0x1801, 26, 23, 0],
	[0x1601, 27, 24, 0],
	[0x1401, 28, 25, 0],
	[0x1201, 29, 26, 0],
	[0x1101, 30, 27, 0],
	[0x0ac1, 31, 28, 0],
	[0x09c1, 32, 29, 0],
	[0x08a1, 33, 30, 0],
	[0x0521, 34, 31, 0],
	[0x0441, 35, 32, 0],
	[0x02a1, 36, 33, 0],
	[0x0221, 37, 34, 0],
	[0x0141, 38, 35, 0],
	[0x0111, 39, 36, 0],
	[0x0085, 40, 37, 0],
	[0x0049, 41, 38, 0],
	[0x0025, 42, 39, 0],
	[0x0015, 43, 40, 0],
	[0x0009, 44, 41, 0],
	[0x0005, 45, 42, 0],
	[0x0001, 45, 43, 0],
	[0x5601, 46, 46, 0],
] as const;

/** The MQ encoder (E.2), ending with the 0xFF 0xAC marker as the spec's FLUSH does. */
export class MqEncoder {
	private a = 0x8000;
	private c = 0;
	private ct = 12;
	/** Index 0 is the byte before the data (never 0xFF). */
	private readonly out: number[] = [0];

	encode(contexts: Uint8Array, at: number, bit: number): void {
		let index = contexts[at]! >> 1;
		let mps = contexts[at]! & 1;
		const [qe, nmps, nlps, sw] = QE[index]!;
		this.a -= qe;
		if (bit === mps) {
			if ((this.a & 0x8000) === 0) {
				if (this.a < qe) this.a = qe;
				else this.c += qe;
				index = nmps;
				this.renormalize();
			} else this.c += qe;
		} else {
			if (this.a < qe) this.c += qe;
			else this.a = qe;
			if (sw) mps = 1 - mps;
			index = nlps;
			this.renormalize();
		}
		contexts[at] = (index << 1) | mps;
	}

	private renormalize(): void {
		do {
			this.a <<= 1;
			this.c *= 2;
			this.ct--;
			if (this.ct === 0) this.byteOut();
		} while ((this.a & 0x8000) === 0);
	}

	private byteOut(): void {
		const last = this.out.length - 1;
		if (this.out[last] === 0xff) {
			this.out.push(Math.floor(this.c / 2 ** 20) & 0xff);
			this.c %= 2 ** 20;
			this.ct = 7;
		} else if (this.c < 0x8000000) {
			this.out.push(Math.floor(this.c / 2 ** 19) & 0xff);
			this.c %= 2 ** 19;
			this.ct = 8;
		} else {
			this.out[last]!++;
			if (this.out[last] === 0xff) {
				this.c %= 0x8000000;
				this.out.push(Math.floor(this.c / 2 ** 20) & 0xff);
				this.c %= 2 ** 20;
				this.ct = 7;
			} else {
				this.out.push(Math.floor(this.c / 2 ** 19) & 0xff);
				this.c %= 2 ** 19;
				this.ct = 8;
			}
		}
	}

	finish(): Uint8Array {
		const temp = this.c + this.a;
		this.c = this.c - (this.c % 0x10000) + 0xffff;
		if (this.c >= temp) this.c -= 0x8000;
		this.c *= 2 ** this.ct;
		this.byteOut();
		this.c *= 2 ** this.ct;
		this.byteOut();
		if (this.out.at(-1) !== 0xff) this.out.push(0xff);
		this.out.push(0xac);
		return Uint8Array.from(this.out.slice(1));
	}
}

/** Encodes an integer (A.2), or OOB for null. */
export function encodeInteger(
	encoder: MqEncoder,
	contexts: Uint8Array,
	value: number | null,
): void {
	let prev = 1;
	const bit = (b: number) => {
		encoder.encode(contexts, prev, b);
		prev = prev < 256 ? (prev << 1) | b : (((prev << 1) | b) & 511) | 256;
	};
	const bits = (v: number, n: number) => {
		for (let i = n - 1; i >= 0; i--) bit(Math.floor(v / 2 ** i) & 1);
	};
	if (value === null) {
		bit(1);
		bit(0);
		bits(0, 2);
		return;
	}
	bit(value < 0 ? 1 : 0);
	const v = Math.abs(value);
	const ranges: [number, number][] = [
		[0, 2],
		[4, 4],
		[20, 6],
		[84, 8],
		[340, 12],
		[4436, 32],
	];
	const index = ranges.findIndex(([low, n], i) => i === 5 || v < low + 2 ** n);
	for (let i = 0; i < index; i++) bit(1);
	if (index < 5) bit(0);
	const [low, n] = ranges[index]!;
	bits(v - low, n);
}

export const u32 = (value: number) => [
	(value >>> 24) & 0xff,
	(value >>> 16) & 0xff,
	(value >>> 8) & 0xff,
	value & 0xff,
];

/** A segment header (7.2) and its data; `length` overrides the data length (0xffffffff). */
export function segment(
	number: number,
	type: number,
	data: ArrayLike<number>,
	options: { referred?: number[]; page?: number; length?: number } = {},
): number[] {
	const referred = options.referred ?? [];
	const size = number <= 256 ? 1 : number <= 65_536 ? 2 : 4;
	return [
		...u32(number),
		type,
		referred.length << 5,
		...referred.flatMap((ref) => u32(ref).slice(4 - size)),
		options.page ?? 1,
		...u32(options.length ?? data.length),
		...Array.from(data),
	];
}

/** Page information (7.4.8) data. */
export const pageInfo = (width: number, height: number, flags = 0, striping = 0): number[] => [
	...u32(width),
	...u32(height),
	...u32(0),
	...u32(0),
	flags,
	striping >> 8,
	striping & 0xff,
];

/** Region segment information (7.4.1) data. */
export const regionInfo = (width: number, height: number, x = 0, y = 0, operator = 0): number[] => [
	...u32(width),
	...u32(height),
	...u32(x),
	...u32(y),
	operator,
];

export const atBytes = (at: [number, number][]): number[] =>
	at.flatMap(([x, y]) => [x & 0xff, y & 0xff]);

/** Segments joined into an embedded stream. */
export const joinSegments = (...segments: number[][]): Uint8Array =>
	Uint8Array.from(segments.flat());

/** A small test bitmap (rows of '0' and '1', 1 black). */
export const SAMPLE_ROWS = [
	'0000000000000000000000',
	'0111111000011000011110',
	'0111111000011000011110',
	'0110000000111100110011',
	'0110000000111100110000',
	'0111110001100110110000',
	'0111110001100110110000',
	'0110000001111110110011',
	'0110000011000011011110',
	'0110000011000011011110',
	'0000000000000000000000',
	'0000000000000000000000',
	'1010101010101010101010',
];

/** A decoded bitmap as rows of '0' and '1'. */
export const bitmapRows = (bitmap: { width: number; height: number; data: Uint8Array }) =>
	Array.from({ length: bitmap.height }, (_, y) =>
		Array.from(bitmap.data.subarray(y * bitmap.width, (y + 1) * bitmap.width)).join(''),
	);
