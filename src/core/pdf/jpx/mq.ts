/**
 * The MQ arithmetic decoder (ITU-T T.800 Annex C) and the raw bit reader of the arithmetic
 * coding bypass (D.6). Both read one codeword segment and see 0xFF bytes past its end, as the
 * standard's decoders do. Context states are bytes: the Qe table index shifted left by one, with
 * the more probable symbol in the low bit.
 */

// Qe, next index after an MPS, next index after an LPS, and whether an LPS switches the MPS.
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

/** Arithmetic decoder over one segment. */
export class MqDecoder {
	private at = 0;
	private c = 0;
	private a = 0x8000;
	private ct = 0;

	constructor(private readonly data: Uint8Array) {
		this.c = (this.byte(0) ^ 0xff) << 16;
		this.byteIn();
		this.c = this.c * 128;
		this.ct -= 7;
	}

	private byte(at: number): number {
		return at < this.data.length ? this.data[at]! : 0xff;
	}

	// INITDEC, BYTEIN and DECODE of C.3, whose code register holds the complemented bytes.
	private byteIn(): void {
		if (this.byte(this.at) === 0xff) {
			if (this.byte(this.at + 1) > 0x8f) this.ct = 8;
			else {
				this.at++;
				this.c += 0xfe00 - (this.byte(this.at) << 9);
				this.ct = 7;
			}
		} else {
			this.at++;
			this.c += 0xff00 - (this.byte(this.at) << 8);
			this.ct = 8;
		}
	}

	/** Decodes one decision in context `index` of `contexts`. */
	decode(contexts: Uint8Array, index: number): number {
		const state = contexts[index]!;
		const i = state >> 1;
		let mps = state & 1;
		const qe = QE[i]!;
		let d: number;
		let next: number;
		this.a -= qe;
		if (this.c >>> 16 < this.a) {
			if (this.a & 0x8000) return mps;
			// MPS exchange
			if (this.a < qe) {
				d = 1 - mps;
				if (SWITCH.has(i)) mps = d;
				next = NLPS[i]!;
			} else {
				d = mps;
				next = NMPS[i]!;
			}
		} else {
			this.c -= this.a * 65536;
			// LPS exchange
			if (this.a < qe) {
				d = mps;
				next = NMPS[i]!;
			} else {
				d = 1 - mps;
				if (SWITCH.has(i)) mps = d;
				next = NLPS[i]!;
			}
			this.a = qe;
		}
		do {
			if (this.ct === 0) this.byteIn();
			this.a <<= 1;
			this.c = (this.c << 1) >>> 0;
			this.ct--;
		} while (!(this.a & 0x8000));
		contexts[index] = (next << 1) | mps;
		return d;
	}
}

/** Reads the raw (bypassed) bits of a segment, skipping the bit stuffed after each 0xFF. */
export class RawDecoder {
	private at = 0;
	private c = 0;
	private ct = 0;

	constructor(private readonly data: Uint8Array) {}

	bit(): number {
		if (this.ct === 0) {
			const byte = this.at < this.data.length ? this.data[this.at]! : 0xff;
			if (this.c === 0xff) {
				if (byte > 0x8f) {
					this.c = 0xff;
					this.ct = 8;
				} else {
					this.c = byte;
					this.at++;
					this.ct = 7;
				}
			} else {
				this.c = byte;
				this.at++;
				this.ct = 8;
			}
		}
		this.ct--;
		return (this.c >> this.ct) & 1;
	}
}
