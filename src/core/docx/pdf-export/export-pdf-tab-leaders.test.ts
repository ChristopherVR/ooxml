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
