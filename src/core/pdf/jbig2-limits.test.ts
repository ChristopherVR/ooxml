import { describe, expect, it } from 'vitest';
import { decodeJbig2Pages, Jbig2Error } from './jbig2';
import { integerContexts } from './jbig2-arith';
import {
	atBytes,
	bitmapRows as rows,
	encodeInteger,
	joinSegments as stream,
	MqEncoder,
	pageInfo,
	regionInfo,
	SAMPLE_ROWS as ROWS,
	segment,
	u32,
} from './test-support/jbig2-encode';
import { encodeGeneric, encodeGenericInto } from './test-support/jbig2-encode-regions';

const W = ROWS[0]!.length;
const H = ROWS.length;
describe('file organisations and malformed streams', () => {
	const at: [number, number][] = [
		[3, -1],
		[-3, -1],
		[2, -2],
		[-2, -2],
	];
	const parts = [
		{ number: 0, type: 48, data: pageInfo(W, H) },
		{
			number: 1,
			type: 39,
			data: [...regionInfo(W, H), 0, ...atBytes(at), ...encodeGeneric(ROWS, 0, at, false)],
		},
		{ number: 2, type: 49, data: [] },
		{ number: 3, type: 51, data: [] },
	];
	const header = [0x97, 0x4a, 0x42, 0x32, 0x0d, 0x0a, 0x1a, 0x0a];

	it('read the random-access organisation (all headers, then all data)', () => {
		const headers = parts.flatMap((part) =>
			segment(part.number, part.type, [], { length: part.data.length }),
		);
		const file = Uint8Array.from([
			...header,
			0,
			0,
			0,
			0,
			1,
			...headers,
			...parts.flatMap((part) => part.data),
		]);
		expect(rows(decodeJbig2Pages(file)[0]!)).toEqual(ROWS);
		const sequential = Uint8Array.from([
			...header,
			3,
			...parts.flatMap((part) => segment(part.number, part.type, part.data)),
		]);
		expect(rows(decodeJbig2Pages(sequential)[0]!)).toEqual(ROWS);
	});

	it('refuse oversized, truncated and inconsistent data with a reason', () => {
		const big = stream(segment(0, 48, pageInfo(100_000, 100_000)));
		expect(() => decodeJbig2Pages(big)).toThrow(/size limit/);
		const region = stream(segment(1, 39, parts[1]!.data));
		expect(() => decodeJbig2Pages(region)).toThrow(/before the page information/);
		const truncated = stream(
			segment(0, 48, pageInfo(W, H)),
			segment(1, 39, parts[1]!.data),
		).subarray(0, 60);
		expect(() => decodeJbig2Pages(truncated)).toThrow(Jbig2Error);
		const extended = stream(
			segment(0, 48, pageInfo(W, H)),
			segment(1, 39, [...regionInfo(W, H), 16, ...new Array(24).fill(0)]),
		);
		expect(() => decodeJbig2Pages(extended)).toThrow(/extended generic templates/);
		// A symbol dictionary that claims four billion new symbols is refused before decoding.
		const dictionary = (data: number[], count = 1) => [
			0,
			0,
			...atBytes(at),
			0,
			0,
			0,
			1,
			...u32(count),
			...data,
		];
		expect(() => decodeJbig2Pages(stream(segment(0, 0, dictionary([], 0xffffffff))))).toThrow(
			/too many symbols/,
		);
	});

	it('stop a text region with four billion instances at the work limit', () => {
		// One 1 x 1 symbol: height class 1, width 1, its bitmap, OOB, then export it.
		const encoder = new MqEncoder();
		const iadh = integerContexts();
		const iadw = integerContexts();
		const iaex = integerContexts();
		encodeInteger(encoder, iadh, 1);
		encodeInteger(encoder, iadw, 1);
		encodeGenericInto(encoder, new Uint8Array(1 << 16), ['1'], 0, at, false);
		encodeInteger(encoder, iadw, null);
		encodeInteger(encoder, iaex, 0);
		encodeInteger(encoder, iaex, 1);
		const symbols = [0, 0, ...atBytes(at), 0, 0, 0, 1, 0, 0, 0, 1, ...encoder.finish()];
		// Text region: STRIPT 0, one strip (DT 0, FS 0), then instances 1 apart without end.
		const text = new MqEncoder();
		const [dt, fs, ds] = [integerContexts(), integerContexts(), integerContexts()];
		encodeInteger(text, dt, 0);
		encodeInteger(text, dt, 0);
		encodeInteger(text, fs, 0);
		for (let index = 0; index < 5000; index++) encodeInteger(text, ds, 1);
		const coded = text.finish();
		const build = (count: number) =>
			stream(
				segment(0, 48, pageInfo(64, 8)),
				segment(1, 0, symbols),
				segment(2, 6, [...regionInfo(64, 8), 0, 0x10, ...u32(count), ...coded], { referred: [1] }),
			);
		expect(() => decodeJbig2Pages(build(0xffffffff), undefined, { maxWork: 20_000 })).toThrow(
			/too much work/,
		);
		const [page] = decodeJbig2Pages(build(64));
		expect(rows(page!)).toEqual(['1'.repeat(64), ...new Array(7).fill('0'.repeat(64))]);
	});
});
