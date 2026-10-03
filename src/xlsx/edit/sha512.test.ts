import { createHash, randomBytes } from 'node:crypto';
import { sha512 } from './sha512.js';

const hex = (bytes: Uint8Array) => Buffer.from(bytes).toString('hex');

describe('sha512', () => {
	it('matches the FIPS 180-4 "abc" vector', () => {
		expect(hex(sha512(new TextEncoder().encode('abc')))).toBe(
			'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a' +
				'2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
		);
	});

	it('matches node:crypto across every padding boundary', () => {
		for (const length of [0, 1, 55, 111, 112, 113, 127, 128, 129, 239, 240, 255, 256, 1000, 4097]) {
			const data = randomBytes(length);
			expect(hex(sha512(new Uint8Array(data))), `length ${length}`).toBe(
				createHash('sha512').update(data).digest('hex'),
			);
		}
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
