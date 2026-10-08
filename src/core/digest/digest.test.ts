import { digestFunction, digestSync } from './digest';

// Ported from src/core/pptx/core/utils/digests/digest.test.ts: the dispatcher reaches every
// algorithm ECMA-376 lets a password verifier name, synchronously and in any spelling.
const hex = (bytes: Uint8Array | undefined) => Buffer.from(bytes ?? []).toString('hex');
const abc = new TextEncoder().encode('abc');

describe('digestSync', () => {
	it.each([
		['SHA-256', 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad'],
		['MD5', '900150983cd24fb0d6963f7d28e17f72'],
		['md4', 'a448017aaf21d8525fc10ae87aa6729d'],
		['MD2', 'da853b0d3f88d99b30283a69e6ded6bb'],
		['RIPEMD128', 'c14a12199c66e4ba84636b0f69144c77'],
		['ripemd-160', '8eb208f7e05d987a9b044a8e98c6b087f15a0bfc'],
		[
			'whirlpool',
			'4e2448a4c6f486bb16b6562c73b4020bf3043e3a731bce721ae1b303d97e6d4c' +
				'7181eebdb6c57e277d0e34957114cbd6c797fc9d95d8b582d225292076d4eef5',
		],
	])('computes %s', (name, expected) => {
		expect(hex(digestSync(name, abc))).toBe(expected);
	});

	it('has a function for every recognised name and none for others', () => {
		for (const name of ['SHA-1', 'SHA-384', 'SHA-512', 'MD2', 'MD4', 'MD5', 'WHIRLPOOL'])
			expect(digestFunction(name), name).toBeTypeOf('function');
		for (const name of ['SHA3-256', 'BLAKE2B', '', 'constructor', '__proto__'])
			expect(digestFunction(name), name).toBeUndefined();
	});
});
