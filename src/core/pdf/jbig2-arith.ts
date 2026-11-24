/**
 * The JBIG2 arithmetic decoders (ITU-T T.88 Annex E and A.2, A.3): the MQ decoder, the integer
 * decoders (IADH, IADW, IAEX, ...) and the symbol ID decoder IAID. A context is one byte: the
 * probability estimate index shifted left once, with the MPS sense in bit 0.
 */

/** Qe, NMPS, NLPS and SWITCH for each of the 47 probability states (Table E.1). */
const QE = [
	0x5601, 0x3401, 0x1801, 0x0ac1, 0x0521, 0x0221, 0x5601, 0x5401, 0x4801, 0x3801, 0x3001, 0x2401,
	0x1c01, 0x1601, 0x5601, 0x5401, 0x5101, 0x4801, 0x3801, 0x3401, 0x3001, 0x2801, 0x2401, 0x2201,
	0x1c01, 0x1801, 0x1601, 0x1401, 0x1201, 0x1101, 0x0ac1, 0x09c1, 0x08a1, 0x0521, 0x0441, 0x02a1,
	0x0221, 0x0141, 0x0111, 0x0085, 0x0049, 0x0025, 0x0015, 0x0009, 0x0005, 0x0001, 0x5601,
];
const NMPS = [
	1, 2, 3, 4, 5, 38, 7, 8, 9, 10, 11, 12, 13, 29, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
	27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 45, 46,
];
const NLPS = [
	1, 6, 9, 12, 29, 33, 6, 14, 14, 14, 17, 18, 20, 21, 14, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22,
	23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 46,
];
const SWITCH = new Set([0, 6, 14]);

/** The MQ decoder (E.3) over `data[start, end)`; past the end it reads 1 bits, as the spec says. */
export class ArithmeticDecoder {
	private bp: number;
	private chigh: number;
	private clow = 0;
	private ct = 0;
	private a: number;
	constructor(
		private readonly data: Uint8Array,
		start: number,
		private readonly end: number,
	) {
		this.bp = start;
		this.chigh = start < end ? data[start]! : 0xff;
		this.byteIn();
		this.chigh = ((this.chigh << 7) & 0xffff) | ((this.clow >> 9) & 0x7f);
		this.clow = (this.clow << 7) & 0xffff;
		this.ct -= 7;
		this.a = 0x8000;
	}

	private byte(at: number): number {
		return at < this.end ? this.data[at]! : 0xff;
	}

	private byteIn(): void {
		if (this.byte(this.bp) === 0xff) {
			if (this.byte(this.bp + 1) > 0x8f) {
				this.clow += 0xff00;
				this.ct = 8;
			} else {
				this.bp++;
				this.clow += this.byte(this.bp) << 9;
				this.ct = 7;
			}
		} else {
			this.bp++;
			this.clow += this.byte(this.bp) << 8;
			this.ct = 8;
		}
		if (this.clow > 0xffff) {
			this.chigh += this.clow >> 16;
			this.clow &= 0xffff;
		}
	}

	/** Decodes one bit with context `contexts[at]` (DECODE, E.3.2). */
	bit(contexts: Uint8Array, at: number): number {
		let index = contexts[at]! >> 1;
		let mps = contexts[at]! & 1;
		const qe = QE[index]!;
		let d: number;
		this.a -= qe;
		if (this.chigh < qe) {
			// LPS exchange.
			if (this.a < qe) {
				this.a = qe;
				d = mps;
				index = NMPS[index]!;
			} else {
				this.a = qe;
				d = 1 ^ mps;
				if (SWITCH.has(index)) mps = d;
				index = NLPS[index]!;
			}
		} else {
			this.chigh -= qe;
			if (this.a & 0x8000) return mps;
			// MPS exchange.
			if (this.a < qe) {
				d = 1 ^ mps;
				if (SWITCH.has(index)) mps = d;
				index = NLPS[index]!;
			} else {
				d = mps;
				index = NMPS[index]!;
			}
		}
		do {
			if (this.ct === 0) this.byteIn();
			this.a <<= 1;
			this.chigh = ((this.chigh << 1) & 0xffff) | ((this.clow >> 15) & 1);
			this.clow = (this.clow << 1) & 0xffff;
			this.ct--;
		} while ((this.a & 0x8000) === 0);
		contexts[at] = (index << 1) | mps;
		return d;
	}
}

/** The contexts of one integer decoder (IADH, IADW, ...). */
export const integerContexts = (): Uint8Array => new Uint8Array(512);

/** Decodes an integer (A.2); null is OOB. */
export function decodeInteger(decoder: ArithmeticDecoder, contexts: Uint8Array): number | null {
	let prev = 1;
	const bit = (): number => {
		const value = decoder.bit(contexts, prev);
		prev = prev < 256 ? (prev << 1) | value : (((prev << 1) | value) & 511) | 256;
		return value;
	};
	const bits = (count: number): number => {
		let value = 0;
		for (let index = 0; index < count; index++) value = value * 2 + bit();
		return value;
	};
	const sign = bit();
	const value = !bit()
		? bits(2)
		: !bit()
			? bits(4) + 4
			: !bit()
				? bits(6) + 20
				: !bit()
					? bits(8) + 84
					: !bit()
						? bits(12) + 340
						: bits(32) + 4436;
	if (!sign) return value;
	return value > 0 ? -value : null;
}

/** The contexts of IAID for symbol codes of `length` bits. */
export const idContexts = (length: number): Uint8Array => new Uint8Array(2 << length);

/** Decodes a symbol ID of `length` bits (A.3). */
export function decodeId(decoder: ArithmeticDecoder, contexts: Uint8Array, length: number): number {
	let prev = 1;
	for (let index = 0; index < length; index++) prev = (prev << 1) | decoder.bit(contexts, prev);
	return prev - (1 << length);
}
