import { describe, expect, it } from 'vitest';
import {
	ArithmeticDecoder,
	decodeId,
	decodeInteger,
	idContexts,
	integerContexts,
} from './jbig2-arith';
import { encodeInteger, MqEncoder } from './test-support/jbig2-encode';

// The arithmetic coder test sequence of ITU-T T.88 Annex H.2 (one context throughout).
const DECODED = '000200510000 00c0 0352872aaaaaaaaa82c02000fcd79ef6bf7fed904f46a3bf';
const ENCODED = '84c73bfce1a143040220000041 0dbb86f4317fff88ff37471adb6adfffac';
const hex = (text: string) => Uint8Array.from(Buffer.from(text.replace(/ /g, ''), 'hex'));

describe('the MQ coder', () => {
	it('decodes the T.88 Annex H.2 test sequence', () => {
		const data = hex(ENCODED);
		const decoder = new ArithmeticDecoder(data, 0, data.length);
		const contexts = new Uint8Array(1);
		const out = new Uint8Array(32);
		for (let index = 0; index < 256; index++)
			out[index >> 3] = out[index >> 3]! | (decoder.bit(contexts, 0) << (7 - (index & 7)));
		expect(Buffer.from(out).toString('hex')).toBe(Buffer.from(hex(DECODED)).toString('hex'));
	});

	it('the test encoder reproduces the test sequence', () => {
		const bits = hex(DECODED);
		const encoder = new MqEncoder();
		const contexts = new Uint8Array(1);
		for (let index = 0; index < 256; index++)
			encoder.encode(contexts, 0, (bits[index >> 3]! >> (7 - (index & 7))) & 1);
		expect(Buffer.from(encoder.finish()).toString('hex')).toBe(
			Buffer.from(hex(ENCODED)).toString('hex'),
		);
	});

	it('round-trips integers (with OOB) and symbol IDs', () => {
		const values = [0, 3, 4, 19, 20, -83, 84, 339, -340, 4435, 4436, 123_456, -2_000_000_000, null];
		const encoder = new MqEncoder();
		const contexts = integerContexts();
		for (const value of values) encodeInteger(encoder, contexts, value);
		const data = encoder.finish();
		const decoder = new ArithmeticDecoder(data, 0, data.length);
		const read = integerContexts();
		expect(values.map(() => decodeInteger(decoder, read))).toEqual(values);
		// Past the end of the data the decoder reads 1 bits and still terminates.
		const empty = new ArithmeticDecoder(new Uint8Array(0), 0, 0);
		const ids = idContexts(5);
		for (let index = 0; index < 100; index++) expect(decodeId(empty, ids, 5)).toBeLessThan(32);
	});
});
