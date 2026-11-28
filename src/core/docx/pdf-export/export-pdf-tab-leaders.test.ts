// @vitest-environment node
import { readFileSync } from 'node:fs';
import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { exportDocxPdf } from './index';
import { referenceFonts } from './__fixtures__/word-reference/reference';

/** The decoded content streams of a PDF (fonts and images left out). */
function content(bytes: Uint8Array): string {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

describe('exportDocxPdf dot tab leaders (#396)', async () => {
	const docx = readFileSync(
		new URL('./__fixtures__/tab-leaders/issue-396-tab-leader-dots.docx', import.meta.url),
	);
	const { model } = await loadDocx(new Uint8Array(docx));
	const { bytes } = await exportDocxPdf(model, {
		fonts: referenceFonts(),
		now: new Date('2026-11-28T00:00:00Z'),
	});
	const ops = content(bytes);
	// A leader is a text object of one repeated glyph (the full stop): its start (px, under the 0.75
	// scale), how many dots it has and its width at the drawn pitch.
	const leaders = [
		...ops.matchAll(/BT \/\w+ [\d.]+ Tf 1 0 0 -1 ([\d.]+) [\d.]+ Tm[^[]*\[([^\]]*)\] TJ ET/g),
	]
		.map(([, x, array]) => ({ x: Number(x), glyphs: array!.match(/<[0-9a-f]{4}>/g)! }))
		.filter(({ glyphs }) => glyphs.length > 20 && new Set(glyphs).size === 1)
		.map(({ x, glyphs }) => ({ x, dots: glyphs.length }));

	it("draws the dots of each leader at Word's advance, as many as fit", () => {
		// Word 16: 110, 110, 151 and 137 periods of 2.76 pt, from 212.57, 212.57, 99.38, 138.02 pt.
		expect(leaders.map(({ dots }) => dots)).toEqual([110, 110, 151, 137]);
		const starts = leaders.map(({ x }) => x * 0.75);
		for (const [index, word] of [212.57, 212.57, 99.38, 138.02].entries())
			expect(Math.abs(starts[index]! - word)).toBeLessThan(0.15);
	});

	it('ends the leaders of successive lines on one grid', () => {
		// Word ends every leader at 516.23 pt, 2.76 pt apart per dot, so the dots line up.
		for (const { x, dots } of leaders)
			expect(Math.abs(x * 0.75 + dots * 2.76 - 516.23)).toBeLessThan(0.15);
	});
});

describe("exportDocxPdf dot leaders on Word's page grid (#432)", () => {
	/** The dot leaders of a fixture: how many dots, and where the first starts (pt). */
	async function dotLeaders(name: string) {
		const docx = readFileSync(new URL(`./__fixtures__/tab-leaders/${name}.docx`, import.meta.url));
		const { model } = await loadDocx(new Uint8Array(docx));
		const { bytes } = await exportDocxPdf(model, {
			fonts: referenceFonts(),
			now: new Date('2026-11-28T00:00:00Z'),
		});
		return [
			...content(bytes).matchAll(
				/BT \/\w+ [\d.]+ Tf 1 0 0 -1 ([\d.]+) [\d.]+ Tm[^[]*\[([^\]]*)\] TJ ET/g,
			),
		]
			.map(([, x, array]) => ({
				x: Number(x) * 0.75,
				dots: array!.match(/<[0-9a-f]+>/g)!.flatMap((hex) => hex.slice(1, -1).match(/.{4}/g)!),
			}))
			.filter(({ x, dots }) => x < 400 && dots.length > 20 && new Set(dots).size === 1)
			.map(({ x, dots }) => ({ x, dots: dots.length }));
	}
	const close = (actual: number[], expected: number[], tolerance = 0.15) =>
		actual.forEach((value, index) =>
			expect(Math.abs(value - expected[index]!), `line ${index + 1}`).toBeLessThan(tolerance),
		);

	// Word 16 (Calibri 11 pt, A4, 2.5 cm margins, right tab with a dot leader at the margin): every
	// dot starts at a multiple of its 2.76 pt advance from the page's left edge, and the cells that
	// lie wholly between the text before and the text after the tab are filled (1 to 30 `i` after).
	it('fills the whole cells before the text after a right tab', async () => {
		const leaders = await dotLeaders('issue-432-right-tab-text-after');
		expect(leaders.map(({ dots }) => dots)).toEqual([
			115, 114, 113, 112, 111, 110, 109, 108, 107, 106, 105, 105, 104, 103, 102, 101, 100, 99, 98,
			97, 96, 95, 94, 94, 93, 92, 91, 90, 89, 88,
		]);
		close(
			leaders.map(({ x }) => x),
			leaders.map(() => 204.29),
		);
	});

	// The same with text of growing length before the tab and a number after it.
	it('starts the dots on the grid cell after the text before the tab', async () => {
		const leaders = await dotLeaders('issue-432-right-tab-text-before');
		expect(leaders.map(({ dots }) => dots)).toEqual([
			152, 148, 145, 141, 137, 133, 130, 127, 125, 120, 115, 111, 109, 104, 101, 98, 95, 91, 88, 85,
			82, 79, 76, 73, 70, 68, 66, 62, 59, 56,
		]);
		close(
			leaders.map(({ x }) => x),
			[
				99.38, 110.42, 118.7, 129.74, 140.78, 151.82, 160.1, 168.38, 173.9, 179.42, 193.25, 204.29,
				209.81, 223.61, 231.89, 240.17, 248.45, 259.49, 267.77, 276.05, 284.33, 292.61, 300.89,
				309.17, 317.47, 322.99, 328.51, 339.55, 347.83, 356.11,
			],
		);
	});

	// The grid is the page's, not the margin's: with a 75 pt left margin the dots stay on it.
	it('keeps the grid on the page when the margins move', async () => {
		const leaders = await dotLeaders('issue-432-wide-margin');
		expect(leaders.map(({ dots }) => dots)).toEqual([
			160, 160, 160, 160, 160, 160, 151, 139, 128, 116, 105, 93,
		]);
		close(
			leaders.map(({ x }) => x),
			[77.3, 77.3, 77.3, 77.3, 77.3, 77.3, 104.9, 135.26, 162.86, 193.25, 220.85, 251.21],
		);
	});
});
