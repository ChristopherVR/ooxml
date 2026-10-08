/**
 * Whirlpool message digest (Barreto & Rijmen, ISO/IEC 10118-3:2004 final
 * revision). The largest algorithm ECMA-376 19.2.1.22 allows a
 * `p:modifyVerifier` to name (512 bits) and, like the RIPEMD variants, one
 * only ever identified by `algorithmName` (CAPI never defined a
 * `cryptAlgorithmSid` for it).
 *
 * An AES-like block cipher `W` run in Miyaguchi-Preneel mode over 10
 * rounds; see `whirlpool-table.ts` for where the T-tables and round
 * constants come from. Ported from RHash's public-domain
 * `librhash/whirlpool.c` reference, with one correction: that reference's
 * `hash[i] = state[0][i]` (before the round loop) then `hash[i] ^=
 * state[0][i]` (after) is a two-step feedforward that is easy to
 * mistranslate into a single `hash[i] ^= stateInitial[i] ^ stateFinal[i]`
 * plus the ORIGINAL `hash[i]` again; the original chaining value must
 * appear in the feedforward exactly once (it is already folded into
 * `stateInitial = block ^ originalHash`), not twice, since a second XOR
 * cancels it out entirely. Multi-block inputs (where the chaining value
 * feeding into block 2+ is non-zero) are the only case that exposes the
 * mistake: single-block vectors like `Whirlpool("")` pass either way
 * because the chaining value starts at zero. Caught and fixed against
 * OpenSSL's legacy-provider `whirlpool` digest before release; see
 * `whirlpool.test.ts` for both single- and multi-block vectors.
 *
 * @module digest/whirlpool
 */
import * as table from './whirlpool-table';

// Local bindings: a module runner that turns imports into namespace getters (Vitest's) would
// otherwise pay a getter call on every one of the hash's table lookups, ten times slower.
const TABLES_HI = table.TABLES_HI;
const TABLES_LO = table.TABLES_LO;
const ROUND_CONSTANTS_HI = table.ROUND_CONSTANTS_HI;
const ROUND_CONSTANTS_LO = table.ROUND_CONSTANTS_LO;

const BLOCK_BYTES = 64;
const WORDS = 8;
const ROUNDS = 10;

/**
 * A row of eight 64-bit words held as 32-bit halves: `hi[i]` and `lo[i]` are the high and low
 * halves of word `i`.
 */
interface Words {
	hi: Uint32Array;
	lo: Uint32Array;
}

const words = (): Words => ({ hi: new Uint32Array(WORDS), lo: new Uint32Array(WORDS) });

/**
 * The Whirlpool T-table lookup fused with the round's byte permutation: combines byte `t` (from
 * the most significant) of word `(column - t) mod 8` of `src`, looked up in T-table `t`, into
 * word `column` of `out`. Unrolled: this is the whole cost of the hash.
 */
function op(src: Words, column: number, out: Words): void {
	const { hi: h, lo: l } = src;
	const b0 = h[column]! >>> 24;
	const b1 = (h[(column + 7) & 7]! >>> 16) & 0xff;
	const b2 = (h[(column + 6) & 7]! >>> 8) & 0xff;
	const b3 = h[(column + 5) & 7]! & 0xff;
	const b4 = l[(column + 4) & 7]! >>> 24;
	const b5 = (l[(column + 3) & 7]! >>> 16) & 0xff;
	const b6 = (l[(column + 2) & 7]! >>> 8) & 0xff;
	const b7 = l[(column + 1) & 7]! & 0xff;
	out.hi[column] =
		TABLES_HI[b0]! ^
		TABLES_HI[256 + b1]! ^
		TABLES_HI[512 + b2]! ^
		TABLES_HI[768 + b3]! ^
		TABLES_HI[1024 + b4]! ^
		TABLES_HI[1280 + b5]! ^
		TABLES_HI[1536 + b6]! ^
		TABLES_HI[1792 + b7]!;
	out.lo[column] =
		TABLES_LO[b0]! ^
		TABLES_LO[256 + b1]! ^
		TABLES_LO[512 + b2]! ^
		TABLES_LO[768 + b3]! ^
		TABLES_LO[1024 + b4]! ^
		TABLES_LO[1280 + b5]! ^
		TABLES_LO[1536 + b6]! ^
		TABLES_LO[1792 + b7]!;
}

/** Scratch rows reused across blocks so hashing allocates nothing per round. */
const scratch = { key: words(), state: words(), initial: words(), nextKey: words(), next: words() };

/** Process one 64-byte block, updating `hash` (the chaining value) in place. */
function processBlock(hash: Words, block: Words): void {
	const { key, state, initial, nextKey, next } = scratch;
	for (let i = 0; i < WORDS; i++) {
		initial.hi[i] = block.hi[i]! ^ hash.hi[i]!;
		initial.lo[i] = block.lo[i]! ^ hash.lo[i]!;
	}
	key.hi.set(hash.hi);
	key.lo.set(hash.lo);
	state.hi.set(initial.hi);
	state.lo.set(initial.lo);

	for (let round = 0; round < ROUNDS; round++) {
		for (let j = 0; j < WORDS; j++) op(key, j, nextKey);
		nextKey.hi[0] = nextKey.hi[0]! ^ ROUND_CONSTANTS_HI[round]!;
		nextKey.lo[0] = nextKey.lo[0]! ^ ROUND_CONSTANTS_LO[round]!;
		for (let j = 0; j < WORDS; j++) {
			op(state, j, next);
			next.hi[j] = next.hi[j]! ^ nextKey.hi[j]!;
			next.lo[j] = next.lo[j]! ^ nextKey.lo[j]!;
		}
		key.hi.set(nextKey.hi);
		key.lo.set(nextKey.lo);
		state.hi.set(next.hi);
		state.lo.set(next.lo);
	}

	// Miyaguchi-Preneel feedforward: the original chaining value appears
	// exactly once, already folded into `initial`.
	for (let i = 0; i < WORDS; i++) {
		hash.hi[i] = initial.hi[i]! ^ state.hi[i]!;
		hash.lo[i] = initial.lo[i]! ^ state.lo[i]!;
	}
}

/**
 * Pad `message` per ISO/IEC 10118-1: a `0x80` byte, zero bytes up to a
 * 32-byte-aligned boundary, then the bit length as a final big-endian
 * 64-bit integer (messages need billions of exabytes before the full
 * spec's 256-bit length field would matter, so only the low 64 bits are
 * ever populated, matching the reference implementation).
 */
function pad(message: Uint8Array): Uint8Array {
	let total = message.length + 1;
	while (total % BLOCK_BYTES !== 32) {
		total++;
	}
	total += 32;

	const padded = new Uint8Array(total);
	padded.set(message);
	padded[message.length] = 0x80;
	new DataView(padded.buffer).setBigUint64(total - 8, BigInt(message.length) * 8n, false);
	return padded;
}

/** Compute the Whirlpool digest of `message` (64 bytes). */
export function whirlpool(message: Uint8Array): Uint8Array {
	const padded = pad(message);
	const view = new DataView(padded.buffer);
	const hash = words();
	const block = words();

	for (let offset = 0; offset < padded.length; offset += BLOCK_BYTES) {
		for (let i = 0; i < WORDS; i++) {
			block.hi[i] = view.getUint32(offset + i * 8, false);
			block.lo[i] = view.getUint32(offset + i * 8 + 4, false);
		}
		processBlock(hash, block);
	}

	const out = new Uint8Array(BLOCK_BYTES);
	const outView = new DataView(out.buffer);
	for (let i = 0; i < WORDS; i++) {
		outView.setUint32(i * 8, hash.hi[i]!, false);
		outView.setUint32(i * 8 + 4, hash.lo[i]!, false);
	}
	return out;
}
