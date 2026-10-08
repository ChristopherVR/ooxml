/**
 * 32-bit helpers and the block loop shared by the MD4-family digests (MD4, MD5, RIPEMD-128,
 * RIPEMD-160). All four pad the same way (a `0x80` byte, zeros to 56 mod 64, the bit length as
 * a little-endian uint64) and read each 64-byte block as sixteen little-endian 32-bit words, so
 * each algorithm only supplies its initial state and per-block transform.
 *
 * Ported from `@christophervr/ole2`'s `utils/digests/bit-ops` (which the pptx area forwarded to)
 * so the shared `digest` area needs no ole2 dependency.
 */

/** One operation of a round: `[wordIndex, shift, additiveConstant]`. */
export type RoundOp = readonly [wordIndex: number, shift: number, addConst: number];

/** A round's nonlinear function of three registers. */
export type RoundFunction = (x: number, y: number, z: number) => number;

/** Rotate a 32-bit unsigned value left by `n` bits (1-31). */
export function rotl32(x: number, n: number): number {
	return ((x << n) | (x >>> (32 - n))) >>> 0;
}

/** Add 32-bit values with unsigned wraparound. */
export function add32(...values: number[]): number {
	let sum = 0;
	for (const value of values) sum = (sum + value) >>> 0;
	return sum;
}

/** MD4-style padding (RFC 1320 section 3.1 / RFC 1321 section 3.1). */
function padMd4Style(message: Uint8Array): Uint8Array {
	const paddedLength = (message.length + 9 + 63) & ~63;
	const padded = new Uint8Array(paddedLength);
	padded.set(message);
	padded[message.length] = 0x80;
	new DataView(padded.buffer).setBigUint64(paddedLength - 8, BigInt(message.length) * 8n, true);
	return padded;
}

/**
 * One MD4-style round over the registers `r = [a, b, c, d]`: each operation updates the target
 * register as `rotl(target + f(next three) + x[k] + addConst, s)` and moves the target backward
 * by one, the a, d, c, b cycle of RFC 1320 section 3.4. With `feedback` (MD5, RFC 1321 section
 * 3.4) the following register is added after the rotation.
 */
export function applyRotatingRound(
	r: number[],
	ops: readonly RoundOp[],
	f: RoundFunction,
	x: readonly number[],
	feedback = false,
): void {
	let target = 0;
	for (const [k, s, addConst] of ops) {
		const i1 = (target + 1) & 3;
		const next = r[i1]!;
		const rotated = rotl32(
			add32(r[target]!, f(next, r[(target + 2) & 3]!, r[(target + 3) & 3]!), x[k]!, addConst),
			s,
		);
		r[target] = feedback ? add32(next, rotated) : rotated;
		target = (target + 3) & 3;
	}
}

/**
 * Pad `message`, feed each 64-byte block (sixteen little-endian words) through `transform`,
 * which updates `state` in place, and return the final state as little-endian bytes.
 */
export function mdStyleDigest(
	message: Uint8Array,
	initialState: readonly number[],
	transform: (state: number[], words: readonly number[]) => void,
): Uint8Array {
	const padded = padMd4Style(message);
	const view = new DataView(padded.buffer);
	const state = [...initialState];
	const words = new Array<number>(16).fill(0);
	for (let offset = 0; offset < padded.length; offset += 64) {
		for (let i = 0; i < 16; i++) words[i] = view.getUint32(offset + i * 4, true);
		transform(state, words);
	}
	const out = new Uint8Array(state.length * 4);
	const outView = new DataView(out.buffer);
	for (let i = 0; i < state.length; i++) outView.setUint32(i * 4, state[i]!, true);
	return out;
}
