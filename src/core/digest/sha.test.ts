import { createHash, randomBytes } from 'node:crypto';
import { digestSync, sha1, sha256, sha384, sha512 } from './index.js';

const hex = (bytes: Uint8Array | undefined) => Buffer.from(bytes ?? []).toString('hex');
const ascii = (text: string) => new TextEncoder().encode(text);

const DIGESTS = [
	['sha1', sha1],
	['sha256', sha256],
	['sha384', sha384],
	['sha512', sha512],
] as const;

// Every length around the 64- and 128-byte block and padding boundaries, plus multi-block input.
const LENGTHS = [
	0, 1, 55, 56, 57, 63, 64, 65, 111, 112, 113, 119, 120, 127, 128, 129, 239, 240, 255, 256, 1000,
	4097,
];

describe.each(DIGESTS)('%s', (name, digest) => {
	it('matches node:crypto across every padding boundary', () => {
		for (const length of LENGTHS) {
			const data = new Uint8Array(randomBytes(length));
			expect(hex(digest(data)), `length ${length}`).toBe(
				createHash(name).update(data).digest('hex'),
			);
		}
	});
});

describe('FIPS 180-4 vectors', () => {
	const long = 'abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq';
	it('sha1', () => {
		expect(hex(sha1(ascii('abc')))).toBe('a9993e364706816aba3e25717850c26c9cd0d89d');
		expect(hex(sha1(ascii(long)))).toBe('84983e441c3bd26ebaae4aa1f95129e5e54670f1');
	});
	it('sha256', () => {
		expect(hex(sha256(ascii('abc')))).toBe(
			'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
		);
		expect(hex(sha256(ascii(long)))).toBe(
			'248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1',
		);
	});
	it('sha384', () => {
		expect(hex(sha384(ascii('abc')))).toBe(
			'cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed' +
				'8086072ba1e7cc2358baeca134c825a7',
		);
	});
	it('sha512', () => {
		expect(hex(sha512(ascii('abc')))).toBe(
			'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a' +
				'2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
		);
	});
});

describe('digestSync', () => {
	it('accepts the ECMA-376 and Excel spellings', () => {
		const data = ascii('abc');
		for (const name of ['SHA-512', 'SHA512', 'sha_512', 'sha512', 'Sha-512'])
			expect(hex(digestSync(name, data)), name).toBe(hex(sha512(data)));
		for (const name of ['SHA-1', 'SHA1', 'sha1'])
			expect(digestSync(name, data)).toEqual(sha1(data));
		expect(digestSync('sha-256', data)).toEqual(sha256(data));
		expect(digestSync('SHA384', data)).toEqual(sha384(data));
	});

	it('returns undefined for digests it does not compute', () => {
		for (const name of ['MD5', 'WHIRLPOOL', 'RIPEMD-160', 'SHA-3-256', '', 'constructor'])
			expect(digestSync(name, new Uint8Array()), name).toBeUndefined();
	});

	it('runs a 100,000-round password spin quickly', () => {
		const buffer = new Uint8Array(68);
		let value = new Uint8Array(64);
		const started = performance.now();
		for (let i = 0; i < 100_000; i++) {
			buffer.set(value);
			new DataView(buffer.buffer).setUint32(64, i, true);
			value = sha512(buffer);
		}
		expect(performance.now() - started).toBeLessThan(5_000);
	});
});
