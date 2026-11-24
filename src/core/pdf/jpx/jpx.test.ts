import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodeJpx } from './index';

interface Fixtures {
	width: number;
	height: number;
	sources: Record<string, { depth: number; signed: boolean; sub: number[]; data: string }>;
	cases: {
		name: string;
		source: string;
		tool: string;
		lossy: boolean;
		data: string;
		reference?: string;
	}[];
}

const fixtures = JSON.parse(
	readFileSync(new URL('./__fixtures__/jpx.json', import.meta.url), 'utf8'),
) as Fixtures;
const bytes = (base64: string) => new Uint8Array(Buffer.from(base64, 'base64'));

/** The source planes of a fixture, one array of samples per component. */
function planes(name: string): number[][] {
	const source = fixtures.sources[name]!;
	const data = bytes(source.data);
	const out: number[][] = [];
	let at = 0;
	for (const factor of source.sub) {
		const count = Math.ceil(fixtures.width / factor) * Math.ceil(fixtures.height / factor);
		const plane: number[] = [];
		for (let i = 0; i < count; i++) {
			let value = source.depth > 8 ? (data[at]! << 8) | data[at + 1]! : data[at]!;
			at += source.depth > 8 ? 2 : 1;
			if (source.signed && value >= 128) value -= 256;
			plane.push(value);
		}
		out.push(plane);
	}
	return out;
}

describe('decodeJpx', () => {
	for (const fixture of fixtures.cases)
		it(`decodes ${fixture.name} (${fixture.tool})`, () => {
			const image = decodeJpx(bytes(fixture.data));
			const expected = planes(fixture.source);
			expect(image.width).toBe(fixtures.width);
			expect(image.height).toBe(fixtures.height);
			expect(image.components).toHaveLength(expected.length);
			if (!fixture.lossy) {
				image.components.forEach((component, c) => {
					expect([...component.data]).toEqual(expected[c]);
				});
				return;
			}
			const reference = bytes(fixture.reference!);
			const count = image.components.length;
			let worst = 0;
			let error = 0;
			image.components.forEach((component, c) =>
				component.data.forEach((value, i) => {
					worst = Math.max(worst, Math.abs(value - reference[i * count + c]!));
					error += Math.abs(value - expected[c]![i]!);
				}),
			);
			// Within rounding of OpenJPEG's own decoder, and close to the source pixels.
			expect(worst).toBeLessThanOrEqual(1);
			expect(error / (count * fixtures.width * fixtures.height)).toBeLessThan(8);
		});
});
