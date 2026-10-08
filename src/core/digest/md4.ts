/**
 * MD4 message digest (RFC 1320). Web Crypto never implemented it; it is here so a verifier
 * naming `algorithmName="MD4"` (or the CryptoAPI `cryptAlgorithmSid="2"`) can be checked.
 *
 * A pure-TypeScript port of `@christophervr/ole2`'s `utils/digests/md4`, with the round
 * schedules generated from RFC 1320 section 3.4 instead of listed operation by operation.
 */
import { add32, applyRotatingRound, mdStyleDigest, type RoundOp } from './bit-ops';

const INITIAL_STATE = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476];

/** floor(2^30 * sqrt(2)) and floor(2^30 * sqrt(3)), the round 2 and round 3 constants. */
const K2 = 0x5a827999;
const K3 = 0x6ed9eba1;

const f1 = (x: number, y: number, z: number): number => (x & y) | (~x & z);
const f2 = (x: number, y: number, z: number): number => (x & y) | (x & z) | (y & z);
const f3 = (x: number, y: number, z: number): number => x ^ y ^ z;

/** Sixteen operations: word index `order[i]`, shift `shifts[i % 4]`, constant `k`. */
function round(order: readonly number[], shifts: readonly number[], k: number): RoundOp[] {
	return order.map((word, i) => [word, shifts[i % 4]!, k] as const);
}

const SEQUENTIAL = Array.from({ length: 16 }, (_, i) => i);
const COLUMNS = Array.from({ length: 16 }, (_, i) => (i % 4) * 4 + (i >> 2));
const BIT_REVERSED = [0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15];

const ROUND1 = round(SEQUENTIAL, [3, 7, 11, 19], 0);
const ROUND2 = round(COLUMNS, [3, 5, 9, 13], K2);
const ROUND3 = round(BIT_REVERSED, [3, 9, 11, 15], K3);

function transform(state: number[], x: readonly number[]): void {
	const r = [...state];
	applyRotatingRound(r, ROUND1, f1, x);
	applyRotatingRound(r, ROUND2, f2, x);
	applyRotatingRound(r, ROUND3, f3, x);
	for (let i = 0; i < 4; i++) state[i] = add32(state[i]!, r[i]!);
}

/** Compute the MD4 digest of `message` (16 bytes). */
export function md4(message: Uint8Array): Uint8Array {
	return mdStyleDigest(message, INITIAL_STATE, transform);
}
