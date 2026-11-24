/**
 * The standard Huffman tables B.1 to B.15 of ITU-T T.88 Annex B.5, as table lines.
 */
import { Jbig2Error } from './jbig2-bitmap';

/** A table line: prefix length, range length, range low, and its kind. */
export interface TableLine {
	prefix: number;
	range: number;
	low: number;
	kind?: 'lower' | 'upper' | 'oob';
}

/**
 * Per table: the range lines as flat (low, prefix length, range length) triples, then the lower
 * range line (low, prefix length), the upper range line and the OOB prefix length (0 for none).
 */
type Spec = [
	triples: number[],
	lower: [number, number] | null,
	upper: [number, number],
	oob: number,
];

// prettier-ignore
const STANDARD: Spec[] = [
	[[0, 1, 4, 16, 2, 8, 272, 3, 16], null, [65808, 3], 0],
	[[0, 1, 0, 1, 2, 0, 2, 3, 0, 3, 4, 3, 11, 5, 6], null, [75, 6], 6],
	[[-256, 8, 8, 0, 1, 0, 1, 2, 0, 2, 3, 0, 3, 4, 3, 11, 5, 6], [-257, 8], [75, 7], 6],
	[[1, 1, 0, 2, 2, 0, 3, 3, 0, 4, 4, 3, 12, 5, 6], null, [76, 5], 0],
	[[-255, 7, 8, 1, 1, 0, 2, 2, 0, 3, 3, 0, 4, 4, 3, 12, 5, 6], [-256, 7], [76, 6], 0],
	[
		[-2048, 5, 10, -1024, 4, 9, -512, 4, 8, -256, 4, 7, -128, 5, 6, -64, 5, 5, -32, 4, 5,
			0, 2, 7, 128, 3, 7, 256, 3, 8, 512, 4, 9, 1024, 4, 10],
		[-2049, 6], [2048, 6], 0,
	],
	[
		[-1024, 4, 9, -512, 3, 8, -256, 4, 7, -128, 5, 6, -64, 5, 5, -32, 4, 5, 0, 4, 5, 32, 5, 5,
			64, 5, 6, 128, 4, 7, 256, 3, 8, 512, 3, 9, 1024, 3, 10],
		[-1025, 5], [2048, 5], 0,
	],
	[
		[-15, 8, 3, -7, 9, 1, -5, 8, 1, -3, 9, 0, -2, 7, 0, -1, 4, 0, 0, 2, 1, 2, 5, 0, 3, 6, 0,
			4, 3, 4, 20, 6, 1, 22, 4, 4, 38, 4, 5, 70, 5, 6, 134, 5, 7, 262, 6, 7, 390, 7, 8, 646, 6, 10],
		[-16, 9], [1670, 9], 2,
	],
	[
		[-31, 8, 4, -15, 9, 2, -11, 8, 2, -7, 9, 1, -5, 7, 1, -3, 4, 1, -1, 3, 1, 1, 3, 1, 3, 5, 1,
			5, 6, 1, 7, 3, 5, 39, 6, 2, 43, 4, 5, 75, 4, 6, 139, 5, 7, 267, 5, 8, 523, 6, 8,
			779, 7, 9, 1291, 6, 11],
		[-32, 9], [3339, 9], 2,
	],
	[
		[-21, 7, 4, -5, 8, 0, -4, 7, 0, -3, 5, 0, -2, 2, 2, 2, 5, 0, 3, 6, 0, 4, 7, 0, 5, 8, 0,
			6, 2, 6, 70, 5, 5, 102, 6, 5, 134, 6, 6, 198, 6, 7, 326, 6, 8, 582, 6, 9, 1094, 6, 10,
			2118, 7, 11],
		[-22, 8], [4166, 8], 2,
	],
	[
		[1, 1, 0, 2, 2, 1, 4, 4, 0, 5, 4, 1, 7, 5, 1, 9, 5, 2, 13, 6, 2, 17, 7, 2, 21, 7, 3,
			29, 7, 4, 45, 7, 5, 77, 7, 6],
		null, [141, 7], 0,
	],
	[
		[1, 1, 0, 2, 2, 0, 3, 3, 1, 5, 5, 0, 6, 5, 1, 8, 6, 1, 10, 7, 0, 11, 7, 1, 13, 7, 2,
			17, 7, 3, 25, 7, 4, 41, 8, 5],
		null, [73, 8], 0,
	],
	[
		[1, 1, 0, 2, 3, 0, 3, 4, 0, 4, 5, 0, 5, 4, 1, 7, 3, 3, 15, 6, 1, 17, 6, 2, 21, 6, 3,
			29, 6, 4, 45, 6, 5, 77, 7, 6],
		null, [141, 7], 0,
	],
	// B.14 has no range lines beyond its five values (an upper line without a code).
	[[-2, 3, 0, -1, 3, 0, 0, 1, 0, 1, 3, 0, 2, 3, 0], null, [3, 0], 0],
	[
		[-24, 7, 4, -8, 6, 2, -4, 5, 1, -2, 4, 0, -1, 3, 0, 0, 1, 0, 1, 3, 0, 2, 4, 0, 3, 5, 1,
			5, 6, 2, 9, 7, 4],
		[-25, 7], [25, 7], 0,
	],
];

/** The lines of standard table B.`number` (an upper line with prefix 0 has no code). */
export function standardLines(number: number): TableLine[] {
	const spec = STANDARD[number - 1];
	if (!spec) throw new Jbig2Error(`There is no standard JBIG2 Huffman table B.${number}.`);
	const [triples, lower, upper, oob] = spec;
	const lines: TableLine[] = [];
	for (let at = 0; at + 2 < triples.length; at += 3)
		lines.push({ low: triples[at]!, prefix: triples[at + 1]!, range: triples[at + 2]! });
	if (lower) lines.push({ low: lower[0], prefix: lower[1], range: 32, kind: 'lower' });
	lines.push({ low: upper[0], prefix: upper[1], range: 32, kind: 'upper' });
	if (oob) lines.push({ low: 0, prefix: oob, range: 0, kind: 'oob' });
	return lines;
}
