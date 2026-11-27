import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';
import { loadDocx } from '../index';
import { createFakeMeasurer } from '../layout/index';
import { describe, expect, it } from 'vitest';
import { exportDocxPdf } from './export-pdf';

const fixtures = join(import.meta.dirname, '../layout/fixtures/word-layout-reports');

async function content(name: string): Promise<string> {
	const { model } = await loadDocx(new Uint8Array(await readFile(join(fixtures, name))));
	const { bytes } = await exportDocxPdf(model, { measurer: createFakeMeasurer() });
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		if (/\/Filter \/FlateDecode/.test(match[1]!) && !/\/Subtype \/Image|\/Length1/.test(match[1]!))
			streams.push(inflateSync(bytes.subarray(start, start + length)).toString('latin1'));
	}
	return streams.join('\n');
}

/** Each filled and stroked path: its colours, width and points, in drawing order. */
function paths(stream: string) {
	return [
		...stream.matchAll(
			/([\d.]+ [\d.]+ [\d.]+) rg\n([\d.]+ [\d.]+ [\d.]+) RG\n([\d.]+) w 0 J 1 j \[\] 0 d\n((?:[^\n]+ [mlc]\n)+)h B\*/g,
		),
	].map(([, fill, stroke, width, body]) => ({
		fill,
		stroke,
		width: Number(width),
		commands: body!
			.trim()
			.split('\n')
			.map((line) => {
				const parts = line.split(' ');
				return [parts.at(-1), ...parts.slice(0, -1).map(Number)] as [string, ...number[]];
			}),
	}));
}

describe('content control tags in PDF', () => {
	it("draws inline tags as Word's arc-sided outlines with the title inside", async () => {
		const stream = await content('measured-inline-content-control-tags.docx');
		const tags = paths(stream).filter((path) => path.stroke === '0.627451 0.745098 0.843137');
		// Three tagged controls at each of four sizes, two tags each.
		expect(tags).toHaveLength(24);
		for (const tag of tags) {
			expect(tag.fill).toBe('0.92549 0.94902 0.968627');
			expect(tag.width).toBeCloseTo((1.08 * 96) / 72, 3);
			expect(tag.commands.map(([op]) => op)).toEqual(['m', 'c', 'l', 'c']);
		}
		const [start, end] = tags;
		// Opening tag: edges bow left, so each curve's controls lie left of its ends.
		const [, sx, sy] = start!.commands[0]!;
		const startCurve = start!.commands[1]!;
		expect(startCurve[1]!).toBeLessThan(sx!);
		expect(startCurve[6]! - sy!).toBeCloseTo((-13.44 * 96) / 72, 2);
		const right = start!.commands[2]!;
		expect(right[1]! - (sx! - (2.28 * 96) / 72)).toBeCloseTo((18.5 * 96) / 72, 1);
		// Closing tag: edges bow right, 9.84 pt wide at its widest.
		const [, ex] = end!.commands[0]!;
		expect(end!.commands[1]![1]!).toBeGreaterThan(ex!);
		expect(end!.commands[2]![1]! - ex!).toBeCloseTo(((9.84 - 2.28) * 96) / 72, 2);
		expect(stream).not.toContain('0.909804 0.941176 0.972549 rg');
	});

	it('marks paragraph, row and cell tags with an arrow and bar', async () => {
		const stream = await content('block-content-control-tags.docx');
		const outlines = paths(stream).filter((path) => path.fill === '0.988235 0.988235 0.988235');
		expect(outlines).toHaveLength(6);
		const arrows = [
			...stream.matchAll(
				/0\.627451 0\.745098 0\.843137 rg\n([\d.]+) w 0 J 1 j \[\] 0 d\n[^\n]+ m\n[^\n]+ l\n[^\n]+ l\nh B\*\n[^\n]+ re\nB\*/g,
			),
		];
		expect(arrows).toHaveLength(6);
		for (const [, width] of arrows) expect(Number(width)).toBeCloseTo((0.14 * 96) / 72, 3);
	});
});
