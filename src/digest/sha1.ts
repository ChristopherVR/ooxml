// Synchronous SHA-1 (FIPS 180-4). Broken for collision resistance; it is here only because
// ECMA-376 password hashes and legacy files still name it.
import { pad64, wordsOut } from './sha256.js';

const W = new Uint32Array(80);
const rotl = (x: number, n: number) => (x << n) | (x >>> (32 - n));

function compress(h: Uint32Array, view: DataView, offset: number): void {
	for (let i = 0; i < 16; i++) W[i] = view.getUint32(offset + i * 4);
	for (let i = 16; i < 80; i++) W[i] = rotl(W[i - 3]! ^ W[i - 8]! ^ W[i - 14]! ^ W[i - 16]!, 1);
	let a = h[0]!,
		b = h[1]!,
		c = h[2]!,
		d = h[3]!,
		e = h[4]!;
	for (let i = 0; i < 80; i++) {
		let f: number;
		let k: number;
		if (i < 20) {
			f = (b & c) | (~b & d);
			k = 0x5a827999;
		} else if (i < 40) {
			f = b ^ c ^ d;
			k = 0x6ed9eba1;
		} else if (i < 60) {
			f = (b & c) | (b & d) | (c & d);
			k = 0x8f1bbcdc;
		} else {
			f = b ^ c ^ d;
			k = 0xca62c1d6;
		}
		const t = (rotl(a, 5) + f + e + k + W[i]!) >>> 0;
		e = d;
		d = c;
		c = rotl(b, 30) >>> 0;
		b = a;
		a = t;
	}
	h[0] = (h[0]! + a) >>> 0;
	h[1] = (h[1]! + b) >>> 0;
	h[2] = (h[2]! + c) >>> 0;
	h[3] = (h[3]! + d) >>> 0;
	h[4] = (h[4]! + e) >>> 0;
}

/** The SHA-1 digest (20 bytes) of `data`. */
export function sha1(data: Uint8Array): Uint8Array<ArrayBuffer> {
	const h = new Uint32Array([0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476, 0xc3d2e1f0]);
	const view = pad64(data);
	for (let offset = 0; offset < view.byteLength; offset += 64) compress(h, view, offset);
	return wordsOut(h);
}
