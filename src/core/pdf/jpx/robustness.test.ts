import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readJp2 } from './boxes';
import { decodeCodestream } from './decode';
import { DEFAULT_LIMITS } from './errors';
import { decodeJpx, JpxError } from './index';

const fixtures = JSON.parse(
	readFileSync(new URL('./__fixtures__/jpx.json', import.meta.url), 'utf8'),
) as { cases: { name: string; data: string }[] };
const file = (name: string) =>
	new Uint8Array(Buffer.from(fixtures.cases.find((item) => item.name === name)!.data, 'base64'));

/** Decodes (without the catch-all of `decodeJpx`), accepting only a result or a JpxError. */
const attempt = (data: Uint8Array, limits = {}) => {
	try {
		return decodeCodestream(readJp2(data).codestream, { ...DEFAULT_LIMITS, ...limits });
	} catch (error) {
		expect(error).toBeInstanceOf(JpxError);
		return error as JpxError;
	}
};

describe('malformed JPEG 2000 data', () => {
	it('keeps what decoded before the data ends', () => {
		const whole = file('gray-bypass-layers');
		const full = decodeJpx(whole);
		const cut = attempt(whole.subarray(0, Math.floor(whole.length * 0.7)));
		expect(cut).not.toBeInstanceOf(JpxError);
		if (cut instanceof JpxError) throw cut;
		const partial = cut.components[0]!.data;
		expect(partial.length).toBe(full.components[0]!.data.length);
		// The early layers give an approximation of the picture.
		let error = 0;
		partial.forEach((value, i) => (error += Math.abs(value - full.components[0]!.data[i]!)));
		expect(error / partial.length).toBeLessThan(20);
		for (let length = 0; length < whole.length; length += 7) attempt(whole.subarray(0, length));
	});

	it('fails cleanly on corrupted bytes', () => {
		let seed = 319;
		const random = () => (seed = (seed * 1103515245 + 12345) >>> 0) / 2 ** 32;
		for (const name of ['rgb-97', 'gray-all-modes', 'rgb-markers', 'yuv420-rpcl-tiles', 'rgba'])
			for (let round = 0; round < 60; round++) {
				const data = file(name).slice();
				for (let flips = 1 + Math.floor(random() * 4); flips > 0; flips--)
					data[Math.floor(random() * data.length)] = Math.floor(random() * 256);
				attempt(data, { maxWork: 5_000_000 });
			}
	});

	it('refuses images over the limits', () => {
		expect(() => decodeJpx(file('rgb-97'), { maxWork: 1000 })).toThrow('too costly');
		expect(() => decodeJpx(file('rgb-97'), { maxSamples: 100 })).toThrow('size limit');
		expect(() => decodeJpx(file('rgb-53-tiles'), { maxTiles: 2 })).toThrow('too many tiles');
		expect(() => decodeJpx(Uint8Array.of(1, 2, 3))).toThrow('not a JPEG 2000 file');
		// A SIZ segment declaring an enormous image.
		const huge = file('gray-53').slice();
		huge.set([0x7f, 0xff, 0xff, 0xff, 0x7f, 0xff, 0xff, 0xff], 8);
		expect(() => decodeJpx(huge)).toThrow(JpxError);
	});
});
